import {
  type AdvancedSearchForm,
  BasicRateLimiter,
  Chapter,
  ChapterProviding,
  CloudflareBypassRequestProviding,
  CloudflareError,
  ContentRating,
  Cookie,
  CookieStorageInterceptor,
  DiscoverSection,
  DiscoverSectionItem,
  DiscoverSectionProviding,
  DiscoverSectionType,
  Extension,
  MangaProviding,
  ChapterDetails,
  Metadata,
  PagedResults,
  Request,
  SearchQuery,
  SearchResultItem,
  SearchResultsProviding,
  SortingOption,
  SourceManga,
  TagSection,
} from "@paperback/types";
import * as cheerio from "cheerio";
import * as htmlparser2 from "htmlparser2";
import { URLBuilder } from "../utils/url-builder/base";
import { NovelFireInterceptor } from "./interceptors";
import { type NovelFireMetadata, type NovelFireResult, type NovelFireSearchMeta, SORTS } from "./model";
import { NovelFireSearchForm } from "./forms/SearchForm";
import type { CheerioAPI } from "cheerio";

const baseUrl = "https://novelfire.net";

const VOID_ELEMENTS =
  "area|base|br|col|embed|hr|img|input|link|meta|param|source|track|wbr";
const VOID_ELEMENT_REGEX = new RegExp(`<(${VOID_ELEMENTS})((?:\\s+[^>]*?)?)\\s*/?>`, "gi");

// Turns every void element into a self-closing tag so the chapter is valid XHTML for Readium.
function fixVoidElements(html: string): string {
  return html.replace(VOID_ELEMENT_REGEX, (_match, tag: string, attributes: string) => {
    const attrs = attributes.trim();
    return attrs ? `<${tag.toLowerCase()} ${attrs} />` : `<${tag.toLowerCase()} />`;
  });
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Accepts "/book/slug", "/novel/slug/chapter-1" or a full URL on any mirror domain
// and returns just "slug" or "slug/chapter-1".
function toId(href: string): string {
  const match = href.match(/\/(?:book|novel)\/([^?#]+)/);
  return (match?.[1] ?? href).replace(/^\/+|\/+$/g, "");
}

// Resolves relative or protocol-less image paths against the site.
function absoluteUrl(path: string | undefined): string {
  if (!path) return "";
  if (/^https?:\/\//i.test(path)) return path;
  if (path.startsWith("//")) return `https:${path}`;
  return path.startsWith("/") ? `${baseUrl}${path}` : `${baseUrl}/${path}`;
}

function mapStatus(label: string): string {
  switch (label.trim().toLowerCase()) {
    case "ongoing":
      return "Ongoing";
    case "completed":
      return "Completed";
    case "hiatus":
      return "Hiatus";
    case "dropped":
    case "cancelled":
      return "Dropped";
    default:
      return "Unknown";
  }
}

const ADULT_GENRES = ["adult", "smut", "ecchi", "hentai", "yaoi", "yuri", "mature"];

function mapContentRating(genres: string[]): ContentRating {
  const lower = genres.map((g) => g.toLowerCase());
  if (lower.some((g) => g === "adult" || g === "smut" || g === "hentai")) return ContentRating.ADULT;
  if (lower.some((g) => ADULT_GENRES.includes(g))) return ContentRating.MATURE;
  return ContentRating.EVERYONE;
}

type NovelFireImplementation = Extension &
  SearchResultsProviding &
  MangaProviding &
  ChapterProviding &
  DiscoverSectionProviding &
  CloudflareBypassRequestProviding;

export class NovelFireExtension implements NovelFireImplementation {
  requestManager = new NovelFireInterceptor("main");
  globalRateLimiter = new BasicRateLimiter("rateLimiter", {
    numberOfRequests: 10,
    bufferInterval: 1,
    ignoreImages: true,
  });
  cookieStorageInterceptor = new CookieStorageInterceptor({
    storage: "stateManager",
  });

  async initialise(): Promise<void> {
    this.requestManager.registerInterceptor();
    this.globalRateLimiter.registerInterceptor();
    this.cookieStorageInterceptor.registerInterceptor();
  }

  async cloudflareBypassCompleted(
    _request: unknown,
    cookies: Cookie[],
    _localStorage: Record<string, string>,
  ): Promise<void> {
    for (const cookie of cookies) {
      if (cookie.expires && cookie.expires.getTime() <= Date.now()) continue;
      if (
        cookie.name.startsWith("cf") ||
        cookie.name.startsWith("_cf") ||
        cookie.name.startsWith("__cf")
      ) {
        this.cookieStorageInterceptor.setCookie(cookie);
      }
    }
  }

  async getDiscoverSections(): Promise<DiscoverSection[]> {
    return [
      {
        id: "popular_section",
        title: "Popular",
        type: DiscoverSectionType.featured,
      },
      {
        id: "updated_section",
        title: "Recently Updated",
        type: DiscoverSectionType.chapterUpdates,
      },
      {
        id: "new_section",
        title: "New Novels",
        type: DiscoverSectionType.simpleCarousel,
      },
    ];
  }

  async getDiscoverSectionItems(
    section: DiscoverSection,
    metadata: Metadata | undefined,
  ): Promise<PagedResults<DiscoverSectionItem>> {
    const meta = metadata as NovelFireMetadata | undefined;
    switch (section.id) {
      case "popular_section":
        return this.getPopularSectionItems(section, meta);
      case "updated_section":
        return this.getUpdatedSectionItems(section, meta);
      case "new_section":
        return this.getNewSectionItems(section, meta);
      default:
        return { items: [] };
    }
  }

  async getSortingOptions(_query: SearchQuery<Metadata>): Promise<SortingOption[]> {
    return SORTS.map((s) => ({ id: s.id, label: s.label }));
  }

  async getAdvancedSearchForm(query: SearchQuery<Metadata>): Promise<AdvancedSearchForm> {
    const meta = (query.metadata as { searchMeta?: NovelFireSearchMeta } | undefined)?.searchMeta;
    return new NovelFireSearchForm(meta);
  }

  async getSearchResults(
    query: SearchQuery<Metadata>,
    metadata: Metadata | undefined,
    sortingOption: SortingOption | undefined,
  ): Promise<PagedResults<SearchResultItem>> {
    const paginationMeta = metadata as { page?: number } | undefined;
    const page = paginationMeta?.page ?? 1;

    if (query.title && query.title.trim() !== "") {
      const searchUrl = `${baseUrl}/ajax/searchLive?keyword=${encodeURIComponent(query.title.trim())}&type=title`;
      const [, data] = await Application.scheduleRequest({ url: searchUrl, method: "GET" });
      const jsonString = Application.arrayBufferToUTF8String(data);
      const result = JSON.parse(jsonString) as NovelFireResult;
      const items: SearchResultItem[] = (result?.data ?? []).map((item) => ({
        mangaId: item.slug,
        title: item.title,
        imageUrl: item.image.startsWith("http") ? item.image : `${baseUrl}/${item.image}`,
        subtitle: `${item.total_chapter} chapters`,
      }));
      return { items };
    }

    const searchMeta = (query.metadata as { searchMeta?: NovelFireSearchMeta } | undefined)?.searchMeta;
    const genre = Object.entries(searchMeta?.genres ?? {}).find(([, s]) => s === "included")?.[0] ?? "genre-all";
    const sort = sortingOption?.id ?? "sort-latest-release";
    const status = searchMeta?.status ?? "status-all";

    const request = {
      url: new URLBuilder(baseUrl)
        .addPath(genre)
        .addPath(sort)
        .addPath(status)
        .addPath("all-novel")
        .addQuery("page", String(page))
        .build(),
      method: "GET",
    };

    const $ = await this.fetchCheerio(request);
    const collectedIds: string[] = [];
    const items: SearchResultItem[] = [];

    $(".novel-item").each((_, el) => {
      const novel = $(el);
      const a = novel.find("a");
      const url = a.attr("href") || "";
      const title = a.attr("title") || a.text().trim();
      const imgurl = novel.find("img").attr("data-src") || novel.find("img").attr("src") || "";
      const coverUrl = imgurl.startsWith("http") ? imgurl : `${baseUrl}${imgurl}`;
      const mangaId = toId(url);
      if (title && mangaId && !collectedIds.includes(mangaId)) {
        collectedIds.push(mangaId);
        items.push({ mangaId, title, imageUrl: coverUrl, subtitle: undefined });
      }
    });

    const hasNextPage = !!$(".pagination .page-item.active + .page-item").length;

    return {
      items,
      metadata: hasNextPage ? ({ page: page + 1 } satisfies NovelFireMetadata) : undefined,
    };
  }

  async getMangaDetails(mangaId: string): Promise<SourceManga> {
    const request = {
      url: new URLBuilder(baseUrl).addPath("book").addPath(mangaId).build(),
      method: "GET",
    };
    const $ = await this.fetchCheerio(request);

    const cover = $(".cover > img").first();
    const title =
      $(".novel-title").first().text().trim() || (cover.attr("alt") ?? "").trim();
    const coverUrl = absoluteUrl(cover.attr("data-src") ?? cover.attr("src"));

    const genres: string[] = [];
    $(".categories .property-item").each((_, el) => {
      const genre = $(el).text().trim();
      if (genre) genres.push(genre);
    });
    const tags: string[] = [];
    $(".tag").each((_, el) => {
      const tag = $(el).text().trim();
      if (tag) tags.push(tag);
    });

    const author = $(".author .property-item > span").first().text().trim();
    const status = mapStatus(
      $(".header-stats .ongoing").first().text() ||
        $(".header-stats .completed").first().text() ||
        "",
    );
    const rating = parseFloat($(".nub").first().text().trim());

    const summary = $(".summary .content").first();
    summary.find(".expand").remove();
    const description = Application.decodeHTMLEntities(summary.text().trim());

    const tagGroups: TagSection[] = [];
    if (genres.length > 0) {
      tagGroups.push({
        id: "genres",
        title: "Genres",
        tags: genres.map((genre) => ({
          id: genre.toLowerCase().replace(/\s+/g, "-"),
          title: genre,
        })),
      });
    }
    if (tags.length > 0) {
      tagGroups.push({
        id: "tags",
        title: "Tags",
        tags: tags.map((tag) => ({
          id: tag.toLowerCase().replace(/\s+/g, "-"),
          title: tag,
        })),
      });
    }

    return {
      mangaId: mangaId,
      mangaInfo: {
        primaryTitle: title,
        secondaryTitles: [],
        thumbnailUrl: coverUrl,
        synopsis: description,
        ...(Number.isFinite(rating) ? { rating } : {}),
        ...(author ? { author } : {}),
        contentRating: mapContentRating(genres),
        contentType: "novel",
        status,
        tagGroups,
        shareUrl: this.getMangaShareUrl(mangaId),
      },
    };
  }

  async getChapters(sourceManga: SourceManga): Promise<Chapter[]> {
    const mangaId = sourceManga.mangaId;
    const chapters: Chapter[] = [];
    let page = 1;

    while (true) {
      const request = {
        url: new URLBuilder(baseUrl)
          .addPath("book")
          .addPath(mangaId)
          .addPath("chapters")
          .addQuery("page", String(page))
          .build(),
        method: "GET",
      };

      const $ = await this.fetchCheerio(request);
      const items = $(".chapter-list li");
      if (items.length === 0) break;

      items.each((_, el) => {
        const a = $(el).find("a").first();
        const href = a.attr("href") ?? "";
        if (!href) return;

        const chapterId = toId(href);
        const title = (a.attr("title") ?? a.text()).trim();
        const parsedNo = parseFloat(a.find(".chapter-no").text().trim());
        const datetime = a.find("time").attr("datetime");
        const publishDate = datetime ? new Date(datetime) : undefined;

        chapters.push({
          chapterId,
          title,
          sourceManga,
          // Fall back to the chapter's position so sorting and read tracking still work
          chapNum: Number.isFinite(parsedNo) ? parsedNo : chapters.length + 1,
          publishDate:
            publishDate && !Number.isNaN(publishDate.getTime()) ? publishDate : undefined,
          volume: 0,
          langCode: "en",
          version: "1",
        });
      });

      const hasNextPage = !!$(".pagination .page-item.active + .page-item").length;
      if (!hasNextPage) break;
      page++;
    }

    return chapters;
  }

  async getChapterDetails(chapter: Chapter): Promise<ChapterDetails> {
    let url = new URLBuilder(baseUrl)
      .addPath("book")
      .addPath(chapter.chapterId)
      .build();

    const trail: string[] = [];

    // The site may answer with a redirect page (meta refresh / JS) pointing to a mirror
    // domain that has the real text, so follow up to 3 hops.
    for (let hop = 0; hop < 3; hop++) {
      const page = await this.fetchPage({ url, method: "GET" });
      const content = this.extractChapterText(page.$);

      if (content !== "") {
        const heading = chapter.title ? `<h2>${escapeHtml(chapter.title)}</h2>` : "";
        // Readium parses chapters as XHTML, so bare void tags like <br> must be self-closed
        const body = fixVoidElements(heading + "\n" + content);
        const html = `<html xmlns="http://www.w3.org/1999/xhtml"><head></head><body>${body}</body></html>`;
        return {
          mangaId: chapter.sourceManga.mangaId,
          id: chapter.chapterId,
          type: "html",
          html,
        };
      }

      const title = page.$("title").first().text().trim().slice(0, 80);
      trail.push(`${url} -> ${page.finalUrl} [HTTP ${page.status}] "${title}"`);

      const next = this.findRedirectTarget(page.$, page.html, page.finalUrl);
      if (!next || next === url) break;
      url = next;
    }

    // Throw instead of returning an empty page so the reader shows an error, not a blank screen
    throw new Error(`NovelFire: no chapter text found. Tried: ${trail.join(" | ")}`);
  }

  extractChapterText($: CheerioAPI): string {
    $(".nf-ads, .ads, .adsbygoogle, script, style, iframe, ins, noscript").remove();

    for (const selector of ["#content", "#chapter-container", ".chapter-content", "#chapter-content", ".chapter-body"]) {
      const node = $(selector).first();
      if (node.length === 0 || node.text().trim() === "") continue;
      return (node.html() ?? "").replace(/&nbsp;/g, " ").trim();
    }
    return "";
  }

  findRedirectTarget($: CheerioAPI, html: string, currentUrl: string): string | undefined {
    const refresh = $("meta[http-equiv]")
      .filter((_, el) => ($(el).attr("http-equiv") ?? "").toLowerCase() === "refresh")
      .attr("content");
    const fromMeta = refresh?.match(/url\s*=\s*['"]?([^'";\s]+)/i)?.[1];
    const fromScript = html.match(
      /(?:window\.|document\.|top\.)?location(?:\.href)?\s*(?:=|\.replace\(|\.assign\()\s*['"]([^'"]+)['"]/,
    )?.[1];

    const target = (fromMeta ?? fromScript)?.replace(/\\\//g, "/");
    if (!target || target.startsWith("#") || target.startsWith("javascript:")) return undefined;
    if (/^https?:\/\//i.test(target)) return target;

    const origin = currentUrl.match(/^(https?:\/\/[^/]+)/i)?.[1] ?? baseUrl;
    return target.startsWith("/") ? `${origin}${target}` : `${origin}/${target}`;
  }

  getMangaShareUrl(mangaId: string): string {
    return `${baseUrl}/book/${mangaId}`;
  }

  private async getPopularSectionItems(
    _section: DiscoverSection,
    metadata: NovelFireMetadata | undefined,
  ): Promise<PagedResults<DiscoverSectionItem>> {
    const page = metadata?.page ?? 1;
    const collectedIds: string[] = metadata?.collectedIds ?? [];

    const request = {
      url: new URLBuilder(baseUrl)
        .addPath("genre-all")
        .addPath("sort-popular")
        .addPath("status-all")
        .addPath("all-novel")
        .addQuery("page", page.toString())
        .build(),
      method: "GET",
    };

    const $ = await this.fetchCheerio(request);
    const items: DiscoverSectionItem[] = [];

    $(".novel-item").each((_, el) => {
      const novel = $(el);
      const a = novel.find("a");
      const url = a.attr("href") || "";
      const title = a.attr("title") || "";
      const imgurl =
        novel.find("img").attr("data-src") ||
        novel.find("img").attr("src") ||
        "";
      const coverUrl = `${baseUrl}${imgurl}`;
      const rating =
        parseFloat(novel.find(".badge._br").text().trim()) || undefined;
      const chapters =
        parseInt(novel.find(".novel-stats").text().replace(/\D/g, "")) ||
        undefined;

      const mangaId = toId(url);

      if (title && mangaId && !collectedIds.includes(mangaId)) {
        collectedIds.push(mangaId);
        items.push({
          type: "featuredCarouselItem",
          mangaId: mangaId,
          imageUrl: coverUrl,
          title: title,
          supertitle: rating ? `⭐ ${rating}` : undefined,
        });
      }
    });

    const hasNextPage = !!$(".pagination .page-item.active + .page-item").length;

    return {
      items,
      metadata: hasNextPage ? { page: page + 1, collectedIds } satisfies NovelFireMetadata : undefined,
    };
  }

  private async getUpdatedSectionItems(
    _section: DiscoverSection,
    metadata: NovelFireMetadata | undefined,
  ): Promise<PagedResults<DiscoverSectionItem>> {
    const page = metadata?.page ?? 1;
    const collectedIds: string[] = metadata?.collectedIds ?? [];

    const request = {
      url: new URLBuilder(baseUrl)
        .addPath("genre-all")
        .addPath("sort-latest-release")
        .addPath("status-all")
        .addPath("all-novel")
        .addQuery("page", page.toString())
        .build(),
      method: "GET",
    };

    const $ = await this.fetchCheerio(request);
    const items: DiscoverSectionItem[] = [];

    $(".novel-item").each((_, el) => {
      const novel = $(el);
      const a = novel.find("a");
      const url = a.attr("href") || "";
      const title = a.attr("title") || "";
      const imgurl =
        novel.find("img").attr("data-src") ||
        novel.find("img").attr("src") ||
        "";
      const coverUrl = `${baseUrl}${imgurl}`;
      const chapters =
        parseInt(novel.find(".novel-stats").text().replace(/\D/g, "")) ||
        undefined;

      const mangaId = toId(url);

      if (title && mangaId && !collectedIds.includes(mangaId)) {
        collectedIds.push(mangaId);
        items.push({
          type: "chapterUpdatesCarouselItem",
          mangaId: mangaId,
          chapterId: "",
          imageUrl: coverUrl,
          title: title,
          subtitle: chapters ? `${chapters} chapters` : undefined,
        });
      }
    });

    const hasNextPage = !!$(".pagination .page-item.active + .page-item").length;

    return {
      items,
      metadata: hasNextPage ? { page: page + 1, collectedIds } satisfies NovelFireMetadata : undefined,
    };
  }

  private async getNewSectionItems(
    _section: DiscoverSection,
    metadata: NovelFireMetadata | undefined,
  ): Promise<PagedResults<DiscoverSectionItem>> {
    const page = metadata?.page ?? 1;
    const collectedIds: string[] = metadata?.collectedIds ?? [];

    const request = {
      url: new URLBuilder(baseUrl)
        .addPath("genre-all")
        .addPath("sort-new")
        .addPath("status-all")
        .addPath("all-novel")
        .addQuery("page", page.toString())
        .build(),
      method: "GET",
    };

    const $ = await this.fetchCheerio(request);
    const items: DiscoverSectionItem[] = [];

    $(".novel-item").each((_, el) => {
      const novel = $(el);
      const a = novel.find("a");
      const url = a.attr("href") || "";
      const title = a.attr("title") || "";
      const imgurl =
        novel.find("img").attr("data-src") ||
        novel.find("img").attr("src") ||
        "";
      const coverUrl = `${baseUrl}${imgurl}`;
      const rating =
        parseFloat(novel.find(".badge._br").text().trim()) || undefined;

      const mangaId = toId(url);

      if (title && mangaId && !collectedIds.includes(mangaId)) {
        collectedIds.push(mangaId);
        items.push({
          type: "simpleCarouselItem",
          mangaId: mangaId,
          imageUrl: coverUrl,
          title: title,
          subtitle: rating ? `⭐ ${rating}` : undefined,
        });
      }
    });

    const hasNextPage = !!$(".pagination .page-item.active + .page-item").length;

    return {
      items,
      metadata: hasNextPage ? { page: page + 1, collectedIds } satisfies NovelFireMetadata : undefined,
    };
  }

  checkCloudflareStatus(status: number): void {
    if (status == 503 || status == 403) {
      throw new CloudflareError({ url: baseUrl, method: "GET" });
    }
  }

  async fetchPage(request: Request): Promise<{
    $: CheerioAPI;
    html: string;
    finalUrl: string;
    status: number;
  }> {
    const [response, data] = await Application.scheduleRequest(request);
    this.checkCloudflareStatus(response.status);
    const html = Application.arrayBufferToUTF8String(data);
    const dom = htmlparser2.parseDocument(html);
    return {
      $: cheerio.load(dom),
      html,
      finalUrl: response.url || request.url,
      status: response.status,
    };
  }

  async fetchCheerio(request: Request): Promise<CheerioAPI> {
    const [response, data] = await Application.scheduleRequest(request);
    this.checkCloudflareStatus(response.status);
    const htmlStr = Application.arrayBufferToUTF8String(data);
    const dom = htmlparser2.parseDocument(htmlStr);
    return cheerio.load(dom);
  }
}

export const NovelFire = new NovelFireExtension();
