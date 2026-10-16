const { graphGet } = require("../graph");
const { ok, requireString } = require("../response");
const { formatSearchResult, listText } = require("./formatters");

const tools = [
  {
    name: "search_drive",
    description:
      "Search within a SharePoint document library. Verified documents are shown first, and you can limit results to only verified content.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID to search." },
        query: { type: "string", description: "The keyword or phrase to search for." },
        purpleCheckOnly: {
          type: "boolean",
          description: "Set to true to only return verified or approved documents.",
        },
      },
      required: ["driveId", "query"],
    },
  },
];

const handlers = {
  search_drive: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const query = requireString(params, "query", "a search query");
    const { purpleCheckOnly } = params;

    const data = await graphGet(
      accessToken,
      `/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')`,
      { $expand: "listItem($expand=fields)" }
    );

    let items = (data.value || []).map((item) => {
      const moderationStatus = item.listItem?.fields?.["_ModerationStatus"];
      const isPurpleCheck = moderationStatus === 0 || moderationStatus === "0";
      return { ...item, isPurpleCheck };
    });

    if (purpleCheckOnly) {
      items = items.filter((item) => item.isPurpleCheck);
    } else {
      items.sort((a, b) => {
        if (a.isPurpleCheck === b.isPurpleCheck) return 0;
        return a.isPurpleCheck ? -1 : 1;
      });
    }

    return ok(
      listText(
        purpleCheckOnly
          ? `Verified SharePoint search results for "${query}"`
          : `SharePoint search results for "${query}"`,
        items.length,
        "result",
        items,
        formatSearchResult,
        purpleCheckOnly
          ? `No verified documents matched "${query}".`
          : `No documents matched "${query}".`
      ),
      { driveId, query, purpleCheckOnly: !!purpleCheckOnly, count: items.length, items }
    );
  },
};

module.exports = { tools, handlers };
