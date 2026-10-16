const { graphPost } = require("../graph");
const { ok, optionalPositiveNumber, requireString } = require("../response");
const { formatSearchResult, listText } = require("./formatters");

const tools = [
  {
    name: "search_sharepoint_global",
    description:
      "Search across SharePoint sites and document libraries for files that match a query. Verified documents are shown first, and you can limit results to only verified content.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "The keyword or phrase to search for." },
        size: {
          type: "number",
          description: "Optional number of results to return. Defaults to 25.",
        },
        purpleCheckOnly: {
          type: "boolean",
          description: "Set to true to only return verified or approved documents.",
        },
      },
      required: ["query"],
    },
  },
];

const handlers = {
  search_sharepoint_global: async ({ accessToken, params }) => {
    const query = requireString(params, "query", "a search query");
    const { purpleCheckOnly } = params;
    const size = optionalPositiveNumber(params?.size, "size", {
      integer: true,
      min: 1,
      max: 100,
    }) || 25;

    const effectiveQuery = purpleCheckOnly
      ? `(${query}) AND _ModerationStatus:"0"`
      : query;

    const body = {
      requests: [
        {
          entityTypes: ["driveItem"],
          query: {
            queryString: effectiveQuery,
          },
          from: 0,
          size,
          fields: ["name", "id", "webUrl", "parentReference", "_ModerationStatus"],
        },
      ],
    };

    const data = await graphPost(accessToken, `/search/query`, body);
    const hits = data?.value?.[0]?.hitsContainers?.[0]?.hits || [];

    const results = hits.map((hit) => {
      const moderationStatus =
        hit.resource["_ModerationStatus"] ?? hit.resource.fields?.["_ModerationStatus"];
      const isPurpleCheck = purpleCheckOnly ? true : moderationStatus === 0 || moderationStatus === "0";
      return {
        name: hit.resource.name,
        id: hit.resource.id,
        driveId: hit.resource.parentReference?.driveId,
        webUrl: hit.resource.webUrl,
        siteId: hit.resource.parentReference?.siteId,
        path: hit.resource.parentReference?.path,
        isPurpleCheck,
      };
    });

    results.sort((a, b) => {
      if (a.isPurpleCheck === b.isPurpleCheck) return 0;
      return a.isPurpleCheck ? -1 : 1;
    });

    return ok(
      listText(
        purpleCheckOnly
          ? `Verified SharePoint results for "${query}"`
          : `SharePoint results for "${query}"`,
        results.length,
        "result",
        results,
        formatSearchResult,
        purpleCheckOnly
          ? `No verified documents matched "${query}".`
          : `No SharePoint files matched "${query}".`
      ),
      { query, size, purpleCheckOnly: !!purpleCheckOnly, count: results.length, results }
    );
  },
};

module.exports = { tools, handlers };
