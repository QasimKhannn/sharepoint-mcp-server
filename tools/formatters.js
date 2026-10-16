const { bulletList, formatCount, joinBlocks } = require("../response");

function formatSite(site) {
  return `${site.name || "Untitled site"} — ID: ${site.id || "Unavailable"} — URL: ${
    site.webUrl || "Unavailable"
  }`;
}

function formatDrive(drive) {
  return `${drive.name || "Untitled library"} — ID: ${drive.id || "Unavailable"} — URL: ${
    drive.webUrl || "Unavailable"
  }`;
}

function moderationLabel(item) {
  return item.isPurpleCheck ? "Verified" : "Standard";
}

function itemKind(item) {
  if (item.folder) return "Folder";
  if (item.file) return "File";
  return "Item";
}

function formatDriveItem(item) {
  return `${item.name || "Untitled"} (${itemKind(item)}) — ID: ${item.id || "Unavailable"} — Status: ${moderationLabel(
    item
  )} — URL: ${item.webUrl || "Unavailable"}`;
}

function formatListItem(item) {
  const title =
    item.fields?.Title ||
    item.fields?.title ||
    item.webUrl ||
    item.id ||
    "Untitled item";
  return `${title} — Item ID: ${item.id || "Unavailable"}`;
}

function formatPermission(permission) {
  const grantedTo =
    permission.grantedToV2?.user?.displayName ||
    permission.grantedToV2?.siteUser?.displayName ||
    permission.grantedTo?.user?.displayName ||
    permission.link?.scope ||
    "Unknown recipient";
  const roles =
    Array.isArray(permission.roles) && permission.roles.length
      ? permission.roles.join(", ")
      : "unspecified access";
  return `${grantedTo} — ${roles}`;
}

function formatSearchResult(result) {
  return `${result.name || "Untitled"} — Library ID: ${result.driveId || "Unavailable"} — Site ID: ${
    result.siteId || "Unavailable"
  } — Status: ${result.isPurpleCheck ? "Verified" : "Standard"} — URL: ${result.webUrl || "Unavailable"}`;
}

function listText(title, count, singularLabel, items, formatter, emptyMessage) {
  if (!items.length) {
    return joinBlocks([title, emptyMessage]);
  }

  return joinBlocks([
    title,
    `Found ${formatCount(count, singularLabel)}.`,
    bulletList(items.map(formatter)),
  ]);
}

module.exports = {
  formatSite,
  formatDrive,
  formatDriveItem,
  formatListItem,
  formatPermission,
  formatSearchResult,
  listText,
};
