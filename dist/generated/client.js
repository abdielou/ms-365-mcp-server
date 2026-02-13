import { makeApi, Zodios } from "./hack.js";
import { z } from "zod";
const microsoft_graph_ODataErrors_ErrorDetails = z.object({ code: z.string(), message: z.string(), target: z.string().nullish() }).strict();
const microsoft_graph_ODataErrors_InnerError = z.object({
  "request-id": z.string().describe("Request Id as tracked internally by the service").nullish(),
  "client-request-id": z.string().describe("Client request Id as sent by the client application.").nullish(),
  date: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date when the error occured.").nullish()
}).strict();
const microsoft_graph_ODataErrors_MainError = z.object({
  code: z.string(),
  message: z.string(),
  target: z.string().nullish(),
  details: z.array(microsoft_graph_ODataErrors_ErrorDetails).optional(),
  innerError: microsoft_graph_ODataErrors_InnerError.optional()
}).strict();
const microsoft_graph_ODataErrors_ODataError = z.object({ error: microsoft_graph_ODataErrors_MainError }).strict();
const microsoft_graph_identity = z.object({
  displayName: z.string().describe(
    "The display name of the identity.For drive items, the display name might not always be available or up to date. For example, if a user changes their display name the API might show the new value in a future response, but the items associated with the user don't show up as changed when using delta."
  ).nullish(),
  id: z.string().describe(
    "Unique identifier for the identity or actor. For example, in the access reviews decisions API, this property might record the id of the principal, that is, the group, user, or application that's subject to review."
  ).nullish()
}).strict();
const microsoft_graph_identitySet = z.object({
  application: microsoft_graph_identity.optional(),
  device: microsoft_graph_identity.optional(),
  user: microsoft_graph_identity.optional()
}).strict();
const microsoft_graph_sharepointIds = z.object({
  listId: z.string().describe("The unique identifier (guid) for the item's list in SharePoint.").nullish(),
  listItemId: z.string().describe("An integer identifier for the item within the containing list.").nullish(),
  listItemUniqueId: z.string().describe(
    "The unique identifier (guid) for the item within OneDrive for Business or a SharePoint site."
  ).nullish(),
  siteId: z.string().describe("The unique identifier (guid) for the item's site collection (SPSite).").nullish(),
  siteUrl: z.string().describe("The SharePoint URL for the site that contains the item.").nullish(),
  tenantId: z.string().describe("The unique identifier (guid) for the tenancy.").nullish(),
  webId: z.string().describe("The unique identifier (guid) for the item's site (SPWeb).").nullish()
}).strict();
const microsoft_graph_itemReference = z.object({
  driveId: z.string().describe(
    "Unique identifier of the drive instance that contains the driveItem. Only returned if the item is located in a drive. Read-only."
  ).nullish(),
  driveType: z.string().describe(
    "Identifies the type of drive. Only returned if the item is located in a drive. See drive resource for values."
  ).nullish(),
  id: z.string().describe(
    "Unique identifier of the driveItem in the drive or a listItem in a list. Read-only."
  ).nullish(),
  name: z.string().describe("The name of the item being referenced. Read-only.").nullish(),
  path: z.string().describe("Percent-encoded path that can be used to navigate to the item. Read-only.").nullish(),
  shareId: z.string().describe(
    "A unique identifier for a shared resource that can be accessed via the Shares API."
  ).nullish(),
  sharepointIds: microsoft_graph_sharepointIds.optional(),
  siteId: z.string().describe(
    "For OneDrive for Business and SharePoint, this property represents the ID of the site that contains the parent document library of the driveItem resource or the parent list of the listItem resource. The value is the same as the id property of that site resource. It is an opaque string that consists of three identifiers of the site. For OneDrive, this property is not populated."
  ).nullish()
}).strict();
const microsoft_graph_assignedLicense = z.object({
  disabledPlans: z.array(
    z.string().regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/).uuid()
  ).describe(
    "A collection of the unique identifiers for plans that have been disabled. IDs are available in servicePlans > servicePlanId in the tenant's subscribedSkus or serviceStatus > servicePlanId in the tenant's companySubscription."
  ).optional(),
  skuId: z.string().regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/).uuid().describe(
    "The unique identifier for the SKU. Corresponds to the skuId from subscribedSkus or companySubscription."
  ).nullish()
}).strict();
const microsoft_graph_assignedPlan = z.object({
  assignedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The date and time at which the plan was assigned. The Timestamp type represents date and time information using ISO 8601 format and is always in UTC time. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z."
  ).nullish(),
  capabilityStatus: z.string().describe(
    "Condition of the capability assignment. The possible values are Enabled, Warning, Suspended, Deleted, LockedOut. See a detailed description of each value."
  ).nullish(),
  service: z.string().describe("The name of the service; for example, exchange.").nullish(),
  servicePlanId: z.string().regex(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/).uuid().describe(
    "A GUID that identifies the service plan. For a complete list of GUIDs and their equivalent friendly service names, see Product names and service plan identifiers for licensing."
  ).nullish()
}).strict();
const microsoft_graph_authorizationInfo = z.object({ certificateUserIds: z.array(z.string().nullable()).optional() }).strict();
const microsoft_graph_customSecurityAttributeValue = z.object({}).strict();
const microsoft_graph_user = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  displayName: z.string().describe(
    "The name displayed in the address book for the user. This value is usually the combination of the user's first name, middle initial, and family name. This property is required when a user is created and it can't be cleared during updates. Maximum length is 256 characters. Returned by default. Supports $filter (eq, ne, not , ge, le, in, startsWith, and eq on null values), $orderby, and $search."
  ).nullish(),
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The date and time the user was created, in ISO 8601 format and UTC. The value can't be modified and is automatically populated when the entity is created. Nullable. For on-premises users, the value represents when they were first created in Microsoft Entra ID. Property is null for some users created before June 2018 and on-premises users that were synced to Microsoft Entra ID before June 2018. Read-only. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in)."
  ).nullish(),
  state: z.string().describe(
    "The state or province in the user's address. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values)."
  ).nullish(),
  userPrincipalName: z.string().describe(
    "The user principal name (UPN) of the user. The UPN is an Internet-style sign-in name for the user based on the Internet standard RFC 822. By convention, this value should map to the user's email name. The general format is alias@domain, where the domain must be present in the tenant's collection of verified domains. This property is required when a user is created. The verified domains for the tenant can be accessed from the verifiedDomains property of organization.NOTE: This property can't contain accent characters. Only the following characters are allowed A - Z, a - z, 0 - 9, ' . - _ ! # ^ ~. For the complete list of allowed characters, see username policies. Returned by default. Supports $filter (eq, ne, not, ge, le, in, startsWith, endsWith) and $orderby."
  ).nullish(),
  deletedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "Date and time when this object was deleted. Always null when the object hasn't been deleted."
  ).nullish(),
  aboutMe: z.string().describe(
    "A freeform text entry field for the user to describe themselves. Returned only on $select."
  ).nullish(),
  accountEnabled: z.boolean().describe(
    "true if the account is enabled; otherwise, false. This property is required when a user is created. Returned only on $select. Supports $filter (eq, ne, not, and in)."
  ).nullish(),
  ageGroup: z.string().describe(
    "Sets the age group of the user. Allowed values: null, Minor, NotAdult, and Adult. For more information, see legal age group property definitions. Returned only on $select. Supports $filter (eq, ne, not, and in)."
  ).nullish(),
  assignedLicenses: z.array(microsoft_graph_assignedLicense).describe(
    "The licenses that are assigned to the user, including inherited (group-based) licenses. This property doesn't differentiate between directly assigned and inherited licenses. Use the licenseAssignmentStates property to identify the directly assigned and inherited licenses. Not nullable. Returned only on $select. Supports $filter (eq, not, /$count eq 0, /$count ne 0)."
  ).optional(),
  assignedPlans: z.array(microsoft_graph_assignedPlan).describe(
    "The plans that are assigned to the user. Read-only. Not nullable. Returned only on $select. Supports $filter (eq and not)."
  ).optional(),
  authorizationInfo: microsoft_graph_authorizationInfo.optional(),
  birthday: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The birthday of the user. The Timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014, is 2014-01-01T00:00:00Z. Returned only on $select."
  ).optional(),
  businessPhones: z.array(z.string()).describe(
    "The telephone numbers for the user. NOTE: Although it's a string collection, only one number can be set for this property. Read-only for users synced from the on-premises directory. Returned by default. Supports $filter (eq, not, ge, le, startsWith)."
  ).optional(),
  city: z.string().describe(
    "The city where the user is located. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values)."
  ).nullish(),
  companyName: z.string().describe(
    "The name of the company that the user is associated with. This property can be useful for describing the company that a guest comes from. The maximum length is 64 characters.Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values)."
  ).nullish(),
  consentProvidedForMinor: z.string().describe(
    "Sets whether consent was obtained for minors. Allowed values: null, Granted, Denied, and NotRequired. For more information, see legal age group property definitions. Returned only on $select. Supports $filter (eq, ne, not, and in)."
  ).nullish(),
  country: z.string().describe(
    "The country or region where the user is located; for example, US or UK. Maximum length is 128 characters. Returned only on $select. Supports $filter (eq, ne, not, ge, le, in, startsWith, and eq on null values)."
  ).nullish(),
  creationType: z.string().describe(
    "Indicates whether the user account was created through one of the following methods:  As a regular school or work account (null). As an external account (Invitation). As a local account for an Azure Active Directory B2C tenant (LocalAccount). Through self-service sign-up by an internal user using email verification (EmailVerified). Through self-service sign-up by a guest signing up through a link that is part of a user flow (SelfServiceSignUp). Read-only.Returned only on $select. Supports $filter (eq, ne, not, in)."
  ).nullish(),
  customSecurityAttributes: microsoft_graph_customSecurityAttributeValue.optional(),
  department: z.string().describe(
    "The name of the department in which the user works. Maximum length is 64 characters. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in, and eq on null values)."
  ).nullish(),
  deviceEnrollmentLimit: z.number().gte(-2147483648).lte(2147483647).describe(
    "The limit on the maximum number of devices that the user is permitted to enroll. Allowed values are 5 or 1000."
  ).optional(),
  employeeHireDate: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The date and time when the user was hired or will start work in a future hire. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in)."
  ).nullish(),
  employeeId: z.string().describe(
    "The employee identifier assigned to the user by the organization. The maximum length is 16 characters. Returned only on $select. Supports $filter (eq, ne, not , ge, le, in, startsWith, and eq on null values)."
  ).nullish(),
  employeeLeaveDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The date and time when the user left or will leave the organization. To read this property, the calling app must be assigned the User-LifeCycleInfo.Read.All permission. To write this property, the calling app must be assigned the User.Read.All and User-LifeCycleInfo.ReadWrite.All permissions. To read this property in delegated scenarios, the admin needs at least one of the following Microsoft Entra roles: Lifecycle Workflows Administrator (least privilege), Global Reader. To write this property in delegated scenarios, the admin needs the Global Administrator role. Supports $filter (eq, ne, not , ge, le, in). For more information, see Configure the employeeLeaveDateTime property for a user."
  ).nullish()
}).strict().passthrough();
const microsoft_graph_audio = z.object({
  album: z.string().describe("The title of the album for this audio file.").nullish(),
  albumArtist: z.string().describe("The artist named on the album for the audio file.").nullish(),
  artist: z.string().describe("The performing artist for the audio file.").nullish(),
  bitrate: z.number().describe("Bitrate expressed in kbps.").nullish(),
  composers: z.string().describe("The name of the composer of the audio file.").nullish(),
  copyright: z.string().describe("Copyright information for the audio file.").nullish(),
  disc: z.number().gte(-32768).lte(32767).describe("The number of the disc this audio file came from.").nullish(),
  discCount: z.number().gte(-32768).lte(32767).describe("The total number of discs in this album.").nullish(),
  duration: z.number().describe("Duration of the audio file, expressed in milliseconds").nullish(),
  genre: z.string().describe("The genre of this audio file.").nullish(),
  hasDrm: z.boolean().describe("Indicates if the file is protected with digital rights management.").nullish(),
  isVariableBitrate: z.boolean().describe("Indicates if the file is encoded with a variable bitrate.").nullish(),
  title: z.string().describe("The title of the audio file.").nullish(),
  track: z.number().gte(-2147483648).lte(2147483647).describe("The number of the track on the original disc for this audio file.").nullish(),
  trackCount: z.number().gte(-2147483648).lte(2147483647).describe("The total number of tracks on the original disc for this audio file.").nullish(),
  year: z.number().gte(-2147483648).lte(2147483647).describe("The year the audio file was recorded.").nullish()
}).strict();
const microsoft_graph_album = z.object({
  coverImageItemId: z.string().describe("Unique identifier of the driveItem that is the cover of the album.").nullish()
}).strict();
const microsoft_graph_bundle = z.object({
  album: microsoft_graph_album.optional(),
  childCount: z.number().gte(-2147483648).lte(2147483647).describe("Number of children contained immediately within this container.").nullish()
}).strict();
const microsoft_graph_deleted = z.object({ state: z.string().describe("Represents the state of the deleted item.").nullish() }).strict();
const microsoft_graph_hashes = z.object({
  crc32Hash: z.string().describe("The CRC32 value of the file (if available). Read-only.").nullish(),
  quickXorHash: z.string().describe(
    "A proprietary hash of the file that can be used to determine if the contents of the file change (if available). Read-only."
  ).nullish(),
  sha1Hash: z.string().describe("SHA1 hash for the contents of the file (if available). Read-only.").nullish(),
  sha256Hash: z.string().describe("This property isn't supported. Don't use.").nullish()
}).strict();
const microsoft_graph_file = z.object({
  hashes: microsoft_graph_hashes.optional(),
  mimeType: z.string().describe(
    "The MIME type for the file. This is determined by logic on the server and might not be the value provided when the file was uploaded. Read-only."
  ).nullish(),
  processingMetadata: z.boolean().nullish()
}).strict();
const microsoft_graph_fileSystemInfo = z.object({
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("The UTC date and time the file was created on a client.").nullish(),
  lastAccessedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The UTC date and time the file was last accessed. Available for the recent file list only."
  ).nullish(),
  lastModifiedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("The UTC date and time the file was last modified on a client.").nullish()
}).strict();
const microsoft_graph_folderView = z.object({
  sortBy: z.string().describe("The method by which the folder should be sorted.").nullish(),
  sortOrder: z.string().describe(
    "If true, indicates that items should be sorted in descending order. Otherwise, items should be sorted ascending."
  ).nullish(),
  viewType: z.string().describe("The type of view that should be used to represent the folder.").nullish()
}).strict();
const microsoft_graph_folder = z.object({
  childCount: z.number().gte(-2147483648).lte(2147483647).describe("Number of children contained immediately within this container.").nullish(),
  view: microsoft_graph_folderView.optional()
}).strict();
const microsoft_graph_image = z.object({
  height: z.number().gte(-2147483648).lte(2147483647).describe("Optional. Height of the image, in pixels. Read-only.").nullish(),
  width: z.number().gte(-2147483648).lte(2147483647).describe("Optional. Width of the image, in pixels. Read-only.").nullish()
}).strict();
const microsoft_graph_geoCoordinates = z.object({
  altitude: z.number().describe(
    "Optional. The altitude (height), in feet,  above sea level for the item. Read-only. [Simplified from 3 options]"
  ).nullish(),
  latitude: z.number().describe(
    "Optional. The latitude, in decimal, for the item. Read-only. [Simplified from 3 options]"
  ).nullish(),
  longitude: z.number().describe(
    "Optional. The longitude, in decimal, for the item. Read-only. [Simplified from 3 options]"
  ).nullish()
}).strict();
const microsoft_graph_malware = z.object({
  description: z.string().describe("Contains the virus details for the malware facet.").nullish()
}).strict();
const microsoft_graph_package = z.object({
  type: z.string().describe(
    "A string indicating the type of package. While oneNote is the only currently defined value, you should expect other package types to be returned and handle them accordingly."
  ).nullish()
}).strict();
const microsoft_graph_pendingContentUpdate = z.object({
  queuedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time the pending binary operation was queued in UTC time. Read-only.").nullish()
}).strict();
const microsoft_graph_pendingOperations = z.object({ pendingContentUpdate: microsoft_graph_pendingContentUpdate.optional() }).strict();
const microsoft_graph_driveItem = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  name: z.string().describe("The name of the item. Read-write.").nullish(),
  description: z.string().describe("Provides a user-visible description of the item. Optional.").nullish(),
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time of item creation. Read-only.").optional(),
  lastModifiedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time the item was last modified. Read-only.").optional(),
  content: z.string().describe("The content stream, if the item represents a file.").nullish(),
  createdBy: microsoft_graph_identitySet.optional(),
  eTag: z.string().describe("ETag for the item. Read-only.").nullish(),
  lastModifiedBy: microsoft_graph_identitySet.optional(),
  parentReference: microsoft_graph_itemReference.optional(),
  webUrl: z.string().describe(
    "URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only."
  ).nullish(),
  createdByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
  lastModifiedByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
  audio: microsoft_graph_audio.optional(),
  bundle: microsoft_graph_bundle.optional(),
  cTag: z.string().describe(
    "An eTag for the content of the item. This eTag isn't changed if only the metadata is changed. Note This property isn't returned if the item is a folder. Read-only."
  ).nullish(),
  deleted: microsoft_graph_deleted.optional(),
  file: microsoft_graph_file.optional(),
  fileSystemInfo: microsoft_graph_fileSystemInfo.optional(),
  folder: microsoft_graph_folder.optional(),
  image: microsoft_graph_image.optional(),
  location: microsoft_graph_geoCoordinates.optional(),
  malware: microsoft_graph_malware.optional(),
  package: microsoft_graph_package.optional(),
  pendingOperations: microsoft_graph_pendingOperations.optional()
}).strict().passthrough();
const microsoft_graph_driveItemCollectionResponse = z.object({
  "@odata.count": z.number().int().nullable(),
  "@odata.nextLink": z.string().nullable(),
  value: z.array(microsoft_graph_driveItem)
}).partial().strict();
const microsoft_graph_workbookChartFont = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  bold: z.boolean().describe("Indicates whether the fond is bold.").nullish(),
  color: z.string().describe(
    "The HTML color code representation of the text color. For example #FF0000 represents Red."
  ).nullish(),
  italic: z.boolean().describe("Indicates whether the fond is italic.").nullish(),
  name: z.string().describe("The font name. For example 'Calibri'.").nullish(),
  size: z.number().describe("The size of the font. For example,  11. [Simplified from 3 options]").nullish(),
  underline: z.string().describe(
    "The type of underlining applied to the font. The possible values are: None, Single."
  ).nullish()
}).strict();
const microsoft_graph_workbookChartLineFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  color: z.string().describe("The HTML color code that represents the color of lines in the chart.").nullish()
}).strict();
const microsoft_graph_workbookChartAxisFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  font: microsoft_graph_workbookChartFont.optional(),
  line: microsoft_graph_workbookChartLineFormat.optional()
}).strict();
const microsoft_graph_workbookChartGridlinesFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  line: microsoft_graph_workbookChartLineFormat.optional()
}).strict();
const microsoft_graph_workbookChartGridlines = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  visible: z.boolean().describe("Indicates whether the axis gridlines are visible.").optional(),
  format: microsoft_graph_workbookChartGridlinesFormat.optional()
}).strict();
const microsoft_graph_workbookChartAxisTitleFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  font: microsoft_graph_workbookChartFont.optional()
}).strict();
const microsoft_graph_workbookChartAxisTitle = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  text: z.string().describe("Represents the axis title.").nullish(),
  visible: z.boolean().describe("A Boolean that specifies the visibility of an axis title.").optional(),
  format: microsoft_graph_workbookChartAxisTitleFormat.optional()
}).strict();
const microsoft_graph_workbookChartAxis = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  majorUnit: z.unknown().describe(
    "Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number."
  ).optional(),
  maximum: z.unknown().describe(
    "Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number."
  ).optional(),
  minimum: z.unknown().describe(
    "Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number."
  ).optional(),
  minorUnit: z.unknown().describe(
    "Represents the interval between two minor tick marks. 'Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number."
  ).optional(),
  format: microsoft_graph_workbookChartAxisFormat.optional(),
  majorGridlines: microsoft_graph_workbookChartGridlines.optional(),
  minorGridlines: microsoft_graph_workbookChartGridlines.optional(),
  title: microsoft_graph_workbookChartAxisTitle.optional()
}).strict();
const microsoft_graph_workbookChartAxes = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  categoryAxis: microsoft_graph_workbookChartAxis.optional(),
  seriesAxis: microsoft_graph_workbookChartAxis.optional(),
  valueAxis: microsoft_graph_workbookChartAxis.optional()
}).strict();
const microsoft_graph_workbookChartFill = z.object({ id: z.string().describe("The unique identifier for an entity. Read-only.").optional() }).strict();
const microsoft_graph_workbookChartDataLabelFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional(),
  font: microsoft_graph_workbookChartFont.optional()
}).strict();
const microsoft_graph_workbookChartDataLabels = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  position: z.string().describe(
    "DataLabelPosition value that represents the position of the data label. The possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout."
  ).nullish(),
  separator: z.string().describe("String that represents the separator used for the data labels on a chart.").nullish(),
  showBubbleSize: z.boolean().describe("Boolean value that represents whether the data label bubble size is visible.").nullish(),
  showCategoryName: z.boolean().describe("Boolean value that represents whether the data label category name is visible.").nullish(),
  showLegendKey: z.boolean().describe("Boolean value that represents whether the data label legend key is visible.").nullish(),
  showPercentage: z.boolean().describe("Boolean value that represents whether the data label percentage is visible.").nullish(),
  showSeriesName: z.boolean().describe("Boolean value that represents whether the data label series name is visible.").nullish(),
  showValue: z.boolean().describe("Boolean value that represents whether the data label value is visible.").nullish(),
  format: microsoft_graph_workbookChartDataLabelFormat.optional()
}).strict();
const microsoft_graph_workbookChartAreaFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional(),
  font: microsoft_graph_workbookChartFont.optional()
}).strict();
const microsoft_graph_workbookChartLegendFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional(),
  font: microsoft_graph_workbookChartFont.optional()
}).strict();
const microsoft_graph_workbookChartLegend = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  overlay: z.boolean().describe(
    "Indicates whether the chart legend should overlap with the main body of the chart."
  ).nullish(),
  position: z.string().describe(
    "Represents the position of the legend on the chart. The possible values are: Top, Bottom, Left, Right, Corner, Custom."
  ).nullish(),
  visible: z.boolean().describe("Indicates whether the chart legend is visible.").optional(),
  format: microsoft_graph_workbookChartLegendFormat.optional()
}).strict();
const microsoft_graph_workbookChartSeriesFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional(),
  line: microsoft_graph_workbookChartLineFormat.optional()
}).strict();
const microsoft_graph_workbookChartPointFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional()
}).strict();
const microsoft_graph_workbookChartPoint = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  value: z.unknown().describe("The value of a chart point. Read-only.").optional(),
  format: microsoft_graph_workbookChartPointFormat.optional()
}).strict();
const microsoft_graph_workbookChartSeries = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  name: z.string().describe("The name of a series in a chart.").nullish(),
  format: microsoft_graph_workbookChartSeriesFormat.optional(),
  points: z.array(microsoft_graph_workbookChartPoint).describe("A collection of all points in the series. Read-only.").optional()
}).strict();
const microsoft_graph_workbookChartTitleFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fill: microsoft_graph_workbookChartFill.optional(),
  font: microsoft_graph_workbookChartFont.optional()
}).strict();
const microsoft_graph_workbookChartTitle = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  overlay: z.boolean().describe("Indicates whether the chart title will overlay the chart or not.").nullish(),
  text: z.string().describe("The title text of the chart.").nullish(),
  visible: z.boolean().describe("Indicates whether the chart title is visible.").optional(),
  format: microsoft_graph_workbookChartTitleFormat.optional()
}).strict();
const microsoft_graph_workbookChart = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    height: z.number().describe(
      "Represents the height, in points, of the chart object. [Simplified from 3 options]"
    ).nullish(),
    left: z.number().describe(
      "The distance, in points, from the left side of the chart to the worksheet origin. [Simplified from 3 options]"
    ).nullish(),
    name: z.string().describe("Represents the name of a chart object.").nullish(),
    top: z.number().describe(
      "Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). [Simplified from 3 options]"
    ).nullish(),
    width: z.number().describe(
      "Represents the width, in points, of the chart object. [Simplified from 3 options]"
    ).nullish(),
    axes: microsoft_graph_workbookChartAxes.optional(),
    dataLabels: microsoft_graph_workbookChartDataLabels.optional(),
    format: microsoft_graph_workbookChartAreaFormat.optional(),
    legend: microsoft_graph_workbookChartLegend.optional(),
    series: z.array(microsoft_graph_workbookChartSeries).describe(
      "Represents either a single series or collection of series in the chart. Read-only."
    ).optional(),
    title: microsoft_graph_workbookChartTitle.optional(),
    worksheet: microsoft_graph_workbookWorksheet.optional()
  }).strict()
);
const microsoft_graph_workbookNamedItem = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    comment: z.string().describe("The comment associated with this name.").nullish(),
    name: z.string().describe("The name of the object. Read-only.").nullish(),
    scope: z.string().describe(
      "Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only."
    ).optional(),
    type: z.string().describe(
      "The type of reference is associated with the name. The possible values are: String, Integer, Double, Boolean, Range. Read-only."
    ).nullish(),
    value: z.unknown().describe(
      "The formula that the name is defined to refer to. For example, =Sheet14!$B$2:$H$12 and =4.75. Read-only."
    ).optional(),
    visible: z.boolean().describe("Indicates whether the object is visible.").optional(),
    worksheet: microsoft_graph_workbookWorksheet.optional()
  }).strict()
);
const microsoft_graph_workbookPivotTable = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    name: z.string().describe("The name of the pivot table.").nullish(),
    worksheet: microsoft_graph_workbookWorksheet.optional()
  }).strict()
);
const microsoft_graph_workbookWorksheetProtectionOptions = z.object({
  allowAutoFilter: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow the use of the autofilter feature is enabled."
  ).optional(),
  allowDeleteColumns: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow deleting columns is enabled."
  ).optional(),
  allowDeleteRows: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow deleting rows is enabled."
  ).optional(),
  allowFormatCells: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow formatting cells is enabled."
  ).optional(),
  allowFormatColumns: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow formatting columns is enabled."
  ).optional(),
  allowFormatRows: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow formatting rows is enabled."
  ).optional(),
  allowInsertColumns: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow inserting columns is enabled."
  ).optional(),
  allowInsertHyperlinks: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow inserting hyperlinks is enabled."
  ).optional(),
  allowInsertRows: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow inserting rows is enabled."
  ).optional(),
  allowPivotTables: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow the use of the pivot table feature is enabled."
  ).optional(),
  allowSort: z.boolean().describe(
    "Indicates whether the worksheet protection option to allow the use of the sort feature is enabled."
  ).optional()
}).strict();
const microsoft_graph_workbookWorksheetProtection = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  options: microsoft_graph_workbookWorksheetProtectionOptions.optional(),
  protected: z.boolean().describe("Indicates whether the worksheet is protected.  Read-only.").optional()
}).strict();
const microsoft_graph_workbookIcon = z.object({
  index: z.number().gte(-2147483648).lte(2147483647).describe("The index of the icon in the given set.").optional(),
  set: z.string().describe(
    "The set that the icon is part of. The possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes."
  ).optional()
}).strict();
const microsoft_graph_workbookFilterCriteria = z.object({
  color: z.string().describe("The color applied to the cell.").nullish(),
  criterion1: z.string().describe("A custom criterion.").nullish(),
  criterion2: z.string().describe("A custom criterion.").nullish(),
  dynamicCriteria: z.string().describe("A dynamic formula specified in a custom filter.").optional(),
  filterOn: z.string().describe("Indicates whether a filter is applied to a column.").optional(),
  icon: microsoft_graph_workbookIcon.optional(),
  operator: z.string().describe("An operator in a cell; for example, =, >, <, <=, or <>.").optional(),
  values: z.unknown().describe("The values that appear in the cell.").optional()
}).strict();
const microsoft_graph_workbookFilter = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  criteria: microsoft_graph_workbookFilterCriteria.optional()
}).strict();
const microsoft_graph_workbookTableColumn = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  index: z.number().gte(-2147483648).lte(2147483647).describe(
    "The index of the column within the columns collection of the table. Zero-indexed. Read-only."
  ).optional(),
  name: z.string().describe("The name of the table column.").nullish(),
  values: z.unknown().describe(
    "TRepresents the raw values of the specified range. The data returned could be of type string, number, or a Boolean. Cell that contain an error will return the error string."
  ).optional(),
  filter: microsoft_graph_workbookFilter.optional()
}).strict();
const microsoft_graph_workbookTableRow = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  index: z.number().gte(-2147483648).lte(2147483647).describe(
    "The index of the row within the rows collection of the table. Zero-based. Read-only."
  ).optional(),
  values: z.unknown().describe(
    "The raw values of the specified range. The data returned could be of type string, number, or a Boolean. Any cell that contain an error will return the error string."
  ).optional()
}).strict();
const microsoft_graph_workbookSortField = z.object({
  ascending: z.boolean().describe("Represents whether the sorting is done in an ascending fashion.").optional(),
  color: z.string().describe(
    "Represents the color that is the target of the condition if the sorting is on font or cell color."
  ).nullish(),
  dataOption: z.string().describe(
    "Represents additional sorting options for this field. The possible values are: Normal, TextAsNumber."
  ).optional(),
  icon: microsoft_graph_workbookIcon.optional(),
  key: z.number().gte(-2147483648).lte(2147483647).describe(
    "Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row)."
  ).optional(),
  sortOn: z.string().describe(
    "Represents the type of sorting of this condition. The possible values are: Value, CellColor, FontColor, Icon."
  ).optional()
}).strict();
const microsoft_graph_workbookTableSort = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  fields: z.array(microsoft_graph_workbookSortField).describe("The list of the current conditions last used to sort the table. Read-only.").optional(),
  matchCase: z.boolean().describe("Indicates whether the casing impacted the last sort of the table. Read-only.").optional(),
  method: z.string().describe(
    "The Chinese character ordering method last used to sort the table. The possible values are: PinYin, StrokeCount. Read-only."
  ).optional()
}).strict();
const microsoft_graph_workbookTable = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    highlightFirstColumn: z.boolean().describe("Indicates whether the first column contains special formatting.").optional(),
    highlightLastColumn: z.boolean().describe("Indicates whether the last column contains special formatting.").optional(),
    legacyId: z.string().describe(
      "A legacy identifier used in older Excel clients. The value of the identifier remains the same even when the table is renamed. This property should be interpreted as an opaque string value and shouldn't be parsed to any other type. Read-only."
    ).nullish(),
    name: z.string().describe("The name of the table.").nullish(),
    showBandedColumns: z.boolean().describe(
      "Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier."
    ).optional(),
    showBandedRows: z.boolean().describe(
      "Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier."
    ).optional(),
    showFilterButton: z.boolean().describe(
      "Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row."
    ).optional(),
    showHeaders: z.boolean().describe(
      "Indicates whether the header row is visible or not. This value can be set to show or remove the header row."
    ).optional(),
    showTotals: z.boolean().describe(
      "Indicates whether the total row is visible or not. This value can be set to show or remove the total row."
    ).optional(),
    style: z.string().describe(
      "A constant value that represents the Table style. The possible values are: TableStyleLight1 through TableStyleLight21, TableStyleMedium1 through TableStyleMedium28, TableStyleStyleDark1 through TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified."
    ).nullish(),
    columns: z.array(microsoft_graph_workbookTableColumn).describe("The list of all the columns in the table. Read-only.").optional(),
    rows: z.array(microsoft_graph_workbookTableRow).describe("The list of all the rows in the table. Read-only.").optional(),
    sort: microsoft_graph_workbookTableSort.optional(),
    worksheet: microsoft_graph_workbookWorksheet.optional()
  }).strict()
);
const microsoft_graph_workbookWorksheet = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    name: z.string().describe("The display name of the worksheet.").nullish(),
    position: z.number().gte(-2147483648).lte(2147483647).describe("The zero-based position of the worksheet within the workbook.").optional(),
    visibility: z.string().describe(
      "The visibility of the worksheet. The possible values are: Visible, Hidden, VeryHidden."
    ).optional(),
    charts: z.array(microsoft_graph_workbookChart).describe("The list of charts that are part of the worksheet. Read-only.").optional(),
    names: z.array(microsoft_graph_workbookNamedItem).describe("The list of names that are associated with the worksheet. Read-only.").optional(),
    pivotTables: z.array(microsoft_graph_workbookPivotTable).describe("The list of piot tables that are part of the worksheet.").optional(),
    protection: microsoft_graph_workbookWorksheetProtection.optional(),
    tables: z.array(microsoft_graph_workbookTable).describe("The list of tables that are part of the worksheet. Read-only.").optional()
  }).strict()
);
const microsoft_graph_workbookWorksheetCollectionResponse = z.object({
  "@odata.count": z.number().int().nullable(),
  "@odata.nextLink": z.string().nullable(),
  value: z.array(microsoft_graph_workbookWorksheet)
}).partial().strict();
const create_excel_chart_Body = z.object({ type: z.string(), sourceData: z.unknown(), seriesBy: z.string() }).partial().strict();
const microsoft_graph_workbookRangeBorder = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  color: z.string().describe(
    "The HTML color code that represents the color of the border line. Can either be of the form #RRGGBB, for example 'FFA500', or a named HTML color, for example 'orange'."
  ).nullish(),
  sideIndex: z.string().describe(
    "Indicates the specific side of the border. The possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp. Read-only."
  ).nullish(),
  style: z.string().describe(
    "Indicates the line style for the border. The possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot."
  ).nullish(),
  weight: z.string().describe(
    "The weight of the border around a range. The possible values are: Hairline, Thin, Medium, Thick."
  ).nullish()
}).strict();
const microsoft_graph_workbookRangeFill = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  color: z.string().describe(
    "HTML color code representing the color of the border line. Can either be of the form #RRGGBB, for example 'FFA500', or be a named HTML color, for example 'orange'."
  ).nullish()
}).strict();
const microsoft_graph_workbookRangeFont = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  bold: z.boolean().describe("Inidicates whether the font is bold.").nullish(),
  color: z.string().describe(
    "The HTML color code representation of the text color. For example, #FF0000 represents the color red."
  ).nullish(),
  italic: z.boolean().describe("Inidicates whether the font is italic.").nullish(),
  name: z.string().describe("The font name. For example, 'Calibri'.").nullish(),
  size: z.number().describe("The font size. [Simplified from 3 options]").nullish(),
  underline: z.string().describe(
    "The type of underlining applied to the font. The possible values are: None, Single, Double, SingleAccountant, DoubleAccountant."
  ).nullish()
}).strict();
const microsoft_graph_workbookFormatProtection = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  formulaHidden: z.boolean().describe(
    "Indicates whether Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting."
  ).nullish(),
  locked: z.boolean().describe(
    "Indicates whether Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting."
  ).nullish()
}).strict();
const microsoft_graph_workbookRangeFormat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  columnWidth: z.number().describe(
    "The width of all columns within the range. If the column widths aren't uniform, null will be returned. [Simplified from 3 options]"
  ).nullish(),
  horizontalAlignment: z.string().describe(
    "The horizontal alignment for the specified object. The possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed."
  ).nullish(),
  rowHeight: z.number().describe(
    "The height of all rows in the range. If the row heights aren't uniform null will be returned. [Simplified from 3 options]"
  ).nullish(),
  verticalAlignment: z.string().describe(
    "The vertical alignment for the specified object. The possible values are: Top, Center, Bottom, Justify, Distributed."
  ).nullish(),
  wrapText: z.boolean().describe(
    "Indicates whether Excel wraps the text in the object. A null value indicates that the entire range doesn't have a uniform wrap setting."
  ).nullish(),
  borders: z.array(microsoft_graph_workbookRangeBorder).describe("Collection of border objects that apply to the overall range selected Read-only.").optional(),
  fill: microsoft_graph_workbookRangeFill.optional(),
  font: microsoft_graph_workbookRangeFont.optional(),
  protection: microsoft_graph_workbookFormatProtection.optional()
}).strict();
const microsoft_graph_workbookRangeSort = z.object({ id: z.string().describe("The unique identifier for an entity. Read-only.").optional() }).strict();
const microsoft_graph_workbookRange = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  address: z.string().describe(
    "Represents the range reference in A1-style. Address value contains the Sheet reference (for example, Sheet1!A1:B4). Read-only."
  ).nullish(),
  addressLocal: z.string().describe(
    "Represents range reference for the specified range in the language of the user. Read-only."
  ).nullish(),
  cellCount: z.number().gte(-2147483648).lte(2147483647).describe("Number of cells in the range. Read-only.").optional(),
  columnCount: z.number().gte(-2147483648).lte(2147483647).describe("Represents the total number of columns in the range. Read-only.").optional(),
  columnHidden: z.boolean().describe("Indicates whether all columns of the current range are hidden.").nullish(),
  columnIndex: z.number().gte(-2147483648).lte(2147483647).describe(
    "Represents the column number of the first cell in the range. Zero-indexed. Read-only."
  ).optional(),
  formulas: z.unknown().describe("Represents the formula in A1-style notation.").optional(),
  formulasLocal: z.unknown().describe(
    "Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English '=SUM(A1, 1.5)' formula would become '=SUMME(A1; 1,5)' in German."
  ).optional(),
  formulasR1C1: z.unknown().describe("Represents the formula in R1C1-style notation.").optional(),
  hidden: z.boolean().describe("Represents if all cells of the current range are hidden. Read-only.").nullish(),
  numberFormat: z.unknown().describe("Represents Excel's number format code for the given cell.").optional(),
  rowCount: z.number().gte(-2147483648).lte(2147483647).describe("Returns the total number of rows in the range. Read-only.").optional(),
  rowHidden: z.boolean().describe("Indicates whether all rows of the current range are hidden.").nullish(),
  rowIndex: z.number().gte(-2147483648).lte(2147483647).describe("Returns the row number of the first cell in the range. Zero-indexed. Read-only.").optional(),
  text: z.unknown().describe(
    "Text values of the specified range. The Text value doesn't depend on the cell width. The # sign substitution that happens in Excel UI doesn't affect the text value returned by the API. Read-only."
  ).optional(),
  values: z.unknown().describe(
    "Represents the raw values of the specified range. The data returned can be of type string, number, or a Boolean. Cell that contains an error returns the error string."
  ).optional(),
  valueTypes: z.unknown().describe(
    "Represents the type of data of each cell. The possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error. Read-only."
  ).optional(),
  format: microsoft_graph_workbookRangeFormat.optional(),
  sort: microsoft_graph_workbookRangeSort.optional(),
  worksheet: microsoft_graph_workbookWorksheet.optional()
}).strict();
const microsoft_graph_storagePlanInformation = z.object({
  upgradeAvailable: z.boolean().describe("Indicates whether there are higher storage quota plans available. Read-only.").nullish()
}).strict();
const microsoft_graph_quota = z.object({
  deleted: z.number().describe("Total space consumed by files in the recycle bin, in bytes. Read-only.").nullish(),
  remaining: z.number().describe("Total space remaining before reaching the capacity limit, in bytes. Read-only.").nullish(),
  state: z.string().describe("Enumeration value that indicates the state of the storage space. Read-only.").nullish(),
  storagePlanInformation: microsoft_graph_storagePlanInformation.optional(),
  total: z.number().describe("Total allowed storage space, in bytes. Read-only.").nullish(),
  used: z.number().describe("Total space used, in bytes. Read-only.").nullish()
}).strict();
const microsoft_graph_systemFacet = z.object({}).strict();
const microsoft_graph_listInfo = z.object({
  contentTypesEnabled: z.boolean().describe("If true, indicates that content types are enabled for this list.").nullish(),
  hidden: z.boolean().describe(
    "If true, indicates that the list isn't normally visible in the SharePoint user experience."
  ).nullish(),
  template: z.string().describe(
    "An enumerated value that represents the base list template used in creating the list. Possible values include documentLibrary, genericList, task, survey, announcements, contacts, and more."
  ).nullish()
}).strict();
const microsoft_graph_columnTypes = z.enum([
  "note",
  "text",
  "choice",
  "multichoice",
  "number",
  "currency",
  "dateTime",
  "lookup",
  "boolean",
  "user",
  "url",
  "calculated",
  "location",
  "geolocation",
  "term",
  "multiterm",
  "thumbnail",
  "approvalStatus",
  "unknownFutureValue"
]);
const microsoft_graph_booleanColumn = z.object({}).strict();
const microsoft_graph_calculatedColumn = z.object({
  format: z.string().describe(
    "For dateTime output types, the format of the value. The possible values are: dateOnly or dateTime."
  ).nullish(),
  formula: z.string().describe("The formula used to compute the value for this column.").nullish(),
  outputType: z.string().describe(
    "The output type used to format values in this column. The possible values are: boolean, currency, dateTime, number, or text."
  ).nullish()
}).strict();
const microsoft_graph_choiceColumn = z.object({
  allowTextEntry: z.boolean().describe("If true, allows custom values that aren't in the configured choices.").nullish(),
  choices: z.array(z.string().nullable()).describe("The list of values available for this column.").optional(),
  displayAs: z.string().describe(
    "How the choices are to be presented in the UX. Must be one of checkBoxes, dropDownMenu, or radioButtons"
  ).nullish()
}).strict();
const microsoft_graph_contentApprovalStatusColumn = z.object({}).strict();
const microsoft_graph_currencyColumn = z.object({
  locale: z.string().describe("Specifies the locale from which to infer the currency symbol.").nullish()
}).strict();
const microsoft_graph_dateTimeColumn = z.object({
  displayAs: z.string().describe(
    "How the value should be presented in the UX. Must be one of default, friendly, or standard. See below for more details. If unspecified, treated as default."
  ).nullish(),
  format: z.string().describe(
    "Indicates whether the value should be presented as a date only or a date and time. Must be one of dateOnly or dateTime"
  ).nullish()
}).strict();
const microsoft_graph_defaultColumnValue = z.object({
  formula: z.string().describe("The formula used to compute the default value for the column.").nullish(),
  value: z.string().describe("The direct value to use as the default value for the column.").nullish()
}).strict();
const microsoft_graph_geolocationColumn = z.object({}).strict();
const microsoft_graph_hyperlinkOrPictureColumn = z.object({
  isPicture: z.boolean().describe(
    "Specifies whether the display format used for URL columns is an image or a hyperlink."
  ).nullish()
}).strict();
const microsoft_graph_lookupColumn = z.object({
  allowMultipleValues: z.boolean().describe("Indicates whether multiple values can be selected from the source.").nullish(),
  allowUnlimitedLength: z.boolean().describe(
    "Indicates whether values in the column should be able to exceed the standard limit of 255 characters."
  ).nullish(),
  columnName: z.string().describe("The name of the lookup source column.").nullish(),
  listId: z.string().describe("The unique identifier of the lookup source list.").nullish(),
  primaryLookupColumnId: z.string().describe(
    "If specified, this column is a secondary lookup, pulling an additional field from the list item looked up by the primary lookup. Use the list item looked up by the primary as the source for the column named here."
  ).nullish()
}).strict();
const microsoft_graph_numberColumn = z.object({
  decimalPlaces: z.string().describe(
    "How many decimal places to display. See below for information about the possible values."
  ).nullish(),
  displayAs: z.string().describe(
    "How the value should be presented in the UX. Must be one of number or percentage. If unspecified, treated as number."
  ).nullish(),
  maximum: z.number().describe("The maximum permitted value. [Simplified from 3 options]").nullish(),
  minimum: z.number().describe("The minimum permitted value. [Simplified from 3 options]").nullish()
}).strict();
const microsoft_graph_personOrGroupColumn = z.object({
  allowMultipleSelection: z.boolean().describe("Indicates whether multiple values can be selected from the source.").nullish(),
  chooseFromType: z.string().describe(
    "Whether to allow selection of people only, or people and groups. Must be one of peopleAndGroups or peopleOnly."
  ).nullish(),
  displayAs: z.string().describe("How to display the information about the person or group chosen. See below.").nullish()
}).strict();
const microsoft_graph_columnDefinition = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  name: z.string().describe(
    "The API-facing name of the column as it appears in the fields on a listItem. For the user-facing name, see displayName."
  ).nullish(),
  displayName: z.string().describe("The user-facing name of the column.").nullish(),
  description: z.string().describe("The user-facing description of the column.").nullish(),
  type: microsoft_graph_columnTypes.optional(),
  boolean: microsoft_graph_booleanColumn.optional(),
  calculated: microsoft_graph_calculatedColumn.optional(),
  choice: microsoft_graph_choiceColumn.optional(),
  columnGroup: z.string().describe(
    "For site columns, the name of the group this column belongs to. Helps organize related columns."
  ).nullish(),
  contentApprovalStatus: microsoft_graph_contentApprovalStatusColumn.optional(),
  currency: microsoft_graph_currencyColumn.optional(),
  dateTime: microsoft_graph_dateTimeColumn.optional(),
  defaultValue: microsoft_graph_defaultColumnValue.optional(),
  enforceUniqueValues: z.boolean().describe("If true, no two list items may have the same value for this column.").nullish(),
  geolocation: microsoft_graph_geolocationColumn.optional(),
  hidden: z.boolean().describe("Specifies whether the column is displayed in the user interface.").nullish(),
  hyperlinkOrPicture: microsoft_graph_hyperlinkOrPictureColumn.optional(),
  indexed: z.boolean().describe("Specifies whether the column values can be used for sorting and searching.").nullish(),
  isDeletable: z.boolean().describe("Indicates whether this column can be deleted.").nullish(),
  isReorderable: z.boolean().describe("Indicates whether values in the column can be reordered. Read-only.").nullish(),
  isSealed: z.boolean().describe("Specifies whether the column can be changed.").nullish(),
  lookup: microsoft_graph_lookupColumn.optional(),
  number: microsoft_graph_numberColumn.optional(),
  personOrGroup: microsoft_graph_personOrGroupColumn.optional(),
  propagateChanges: z.boolean().describe(
    "If 'true', changes to this column will be propagated to lists that implement the column."
  ).nullish()
}).strict().passthrough();
const microsoft_graph_contentTypeInfo = z.object({
  id: z.string().describe("The ID of the content type.").nullish(),
  name: z.string().describe("The name of the content type.").nullish()
}).strict();
const microsoft_graph_documentSetContent = z.object({
  contentType: microsoft_graph_contentTypeInfo.optional(),
  fileName: z.string().describe(
    "Name of the file in resource folder that should be added as a default content or a template in the document set."
  ).nullish(),
  folderName: z.string().describe(
    "Folder name in which the file will be placed when a new document set is created in the library."
  ).nullish()
}).strict();
const microsoft_graph_documentSet = z.object({
  allowedContentTypes: z.array(microsoft_graph_contentTypeInfo).describe("Content types allowed in document set.").optional(),
  defaultContents: z.array(microsoft_graph_documentSetContent).describe("Default contents of document set.").optional(),
  propagateWelcomePageChanges: z.boolean().describe("Specifies whether to push welcome page changes to inherited content types.").nullish(),
  shouldPrefixNameToFile: z.boolean().describe("Indicates whether to add the name of the document set to each file name.").nullish(),
  welcomePageUrl: z.string().describe("Welcome page absolute URL.").nullish(),
  sharedColumns: z.array(microsoft_graph_columnDefinition).optional(),
  welcomePageColumns: z.array(microsoft_graph_columnDefinition).optional()
}).strict();
const microsoft_graph_contentTypeOrder = z.object({
  default: z.boolean().describe("Indicates whether this is the default content type").nullish(),
  position: z.number().gte(-2147483648).lte(2147483647).describe("Specifies the position in which the content type appears in the selection UI.").nullish()
}).strict();
const microsoft_graph_columnLink = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  name: z.string().describe("The name of the column  in this content type.").nullish()
}).strict();
const microsoft_graph_contentType = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    associatedHubsUrls: z.array(z.string().nullable()).describe(
      "List of canonical URLs for hub sites with which this content type is associated to. This will contain all hub sites where this content type is queued to be enforced or is already enforced. Enforcing a content type means that the content type is applied to the lists in the enforced sites."
    ).optional(),
    description: z.string().describe("The descriptive text for the item.").nullish(),
    documentSet: microsoft_graph_documentSet.optional(),
    documentTemplate: microsoft_graph_documentSetContent.optional(),
    group: z.string().describe(
      "The name of the group this content type belongs to. Helps organize related content types."
    ).nullish(),
    hidden: z.boolean().describe("Indicates whether the content type is hidden in the list's 'New' menu.").nullish(),
    inheritedFrom: microsoft_graph_itemReference.optional(),
    isBuiltIn: z.boolean().describe("Specifies if a content type is a built-in content type.").nullish(),
    name: z.string().describe("The name of the content type.").nullish(),
    order: microsoft_graph_contentTypeOrder.optional(),
    parentId: z.string().describe("The unique identifier of the content type.").nullish(),
    propagateChanges: z.boolean().describe(
      "If true, any changes made to the content type are pushed to inherited content types and lists that implement the content type."
    ).nullish(),
    readOnly: z.boolean().describe(
      "If true, the content type can't be modified unless this value is first set to false."
    ).nullish(),
    sealed: z.boolean().describe(
      "If true, the content type can't be modified by users or through push-down operations. Only site collection administrators can seal or unseal content types."
    ).nullish(),
    base: microsoft_graph_contentType.optional(),
    baseTypes: z.array(microsoft_graph_contentType).describe("The collection of content types that are ancestors of this content type.").optional(),
    columnLinks: z.array(microsoft_graph_columnLink).describe("The collection of columns that are required by this content type.").optional(),
    columnPositions: z.array(microsoft_graph_columnDefinition).describe("Column order information in a content type.").optional(),
    columns: z.array(microsoft_graph_columnDefinition).describe("The collection of column definitions for this content type.").optional()
  }).strict()
);
const microsoft_graph_itemActionStat = z.object({
  actionCount: z.number().gte(-2147483648).lte(2147483647).describe("The number of times the action took place. Read-only.").nullish(),
  actorCount: z.number().gte(-2147483648).lte(2147483647).describe("The number of distinct actors that performed the action. Read-only.").nullish()
}).strict();
const microsoft_graph_incompleteData = z.object({
  missingDataBeforeDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("The service does not have source data before the specified time.").nullish(),
  wasThrottled: z.boolean().describe("Some data was not recorded due to excessive activity.").nullish()
}).strict();
const microsoft_graph_accessAction = z.object({}).strict();
const microsoft_graph_itemActivity = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  access: microsoft_graph_accessAction.optional(),
  activityDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Details about when the activity took place. Read-only.").nullish(),
  actor: microsoft_graph_identitySet.optional(),
  driveItem: microsoft_graph_driveItem.describe("[Note: Simplified from 45 properties to 25 most common ones]").optional()
}).strict();
const microsoft_graph_itemActivityStat = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  access: microsoft_graph_itemActionStat.optional(),
  create: microsoft_graph_itemActionStat.optional(),
  delete: microsoft_graph_itemActionStat.optional(),
  edit: microsoft_graph_itemActionStat.optional(),
  endDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("When the interval ends. Read-only.").nullish(),
  incompleteData: microsoft_graph_incompleteData.optional(),
  isTrending: z.boolean().describe("Indicates whether the item is 'trending.' Read-only.").nullish(),
  move: microsoft_graph_itemActionStat.optional(),
  startDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("When the interval starts. Read-only.").nullish(),
  activities: z.array(microsoft_graph_itemActivity).describe("Exposes the itemActivities represented in this itemActivityStat resource.").optional()
}).strict();
const microsoft_graph_itemAnalytics = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  allTime: microsoft_graph_itemActivityStat.optional(),
  itemActivityStats: z.array(microsoft_graph_itemActivityStat).optional(),
  lastSevenDays: microsoft_graph_itemActivityStat.optional()
}).strict();
const microsoft_graph_publicationFacet = z.object({
  checkedOutBy: microsoft_graph_identitySet.optional(),
  level: z.string().describe(
    "The state of publication for this document. Either published or checkout. Read-only."
  ).nullish(),
  versionId: z.string().describe(
    "The unique identifier for the version that is visible to the current caller. Read-only."
  ).nullish()
}).strict();
const microsoft_graph_fieldValueSet = z.object({ id: z.string().describe("The unique identifier for an entity. Read-only.").optional() }).strict();
const microsoft_graph_documentSetVersionItem = z.object({
  itemId: z.string().describe("The unique identifier for the item.").nullish(),
  title: z.string().describe("The title of the item.").nullish(),
  versionId: z.string().describe("The version ID of the item.").nullish()
}).strict();
const microsoft_graph_documentSetVersion = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  lastModifiedBy: microsoft_graph_identitySet.optional(),
  lastModifiedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time the version was last modified. Read-only.").nullish(),
  publication: microsoft_graph_publicationFacet.optional(),
  fields: microsoft_graph_fieldValueSet.optional(),
  comment: z.string().describe("Comment about the captured version.").nullish(),
  createdBy: microsoft_graph_identitySet.optional(),
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time when this version was created.").nullish(),
  items: z.array(microsoft_graph_documentSetVersionItem).describe("Items within the document set that are captured as part of this version.").optional(),
  shouldCaptureMinorVersion: z.boolean().describe(
    "If true, minor versions of items are also captured; otherwise, only major versions are captured. The default value is false."
  ).nullish()
}).strict();
const microsoft_graph_listItemVersion = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  lastModifiedBy: microsoft_graph_identitySet.optional(),
  lastModifiedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time the version was last modified. Read-only.").nullish(),
  publication: microsoft_graph_publicationFacet.optional(),
  fields: microsoft_graph_fieldValueSet.optional()
}).strict();
const microsoft_graph_listItem = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  createdBy: microsoft_graph_identitySet.optional(),
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time of item creation. Read-only.").optional(),
  description: z.string().describe("Provides a user-visible description of the item. Optional.").nullish(),
  eTag: z.string().describe("ETag for the item. Read-only.").nullish(),
  lastModifiedBy: microsoft_graph_identitySet.optional(),
  lastModifiedDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe("Date and time the item was last modified. Read-only.").optional(),
  name: z.string().describe("The name of the item. Read-write.").nullish(),
  parentReference: microsoft_graph_itemReference.optional(),
  webUrl: z.string().describe(
    "URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only."
  ).nullish(),
  createdByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
  lastModifiedByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
  contentType: microsoft_graph_contentTypeInfo.optional(),
  deleted: microsoft_graph_deleted.optional(),
  sharepointIds: microsoft_graph_sharepointIds.optional(),
  analytics: microsoft_graph_itemAnalytics.optional(),
  documentSetVersions: z.array(microsoft_graph_documentSetVersion).describe("Version information for a document set version created by a user.").optional(),
  driveItem: microsoft_graph_driveItem.describe("[Note: Simplified from 45 properties to 25 most common ones]").optional(),
  fields: microsoft_graph_fieldValueSet.optional(),
  versions: z.array(microsoft_graph_listItemVersion).describe("The list of previous versions of the list item.").optional()
}).strict();
const microsoft_graph_longRunningOperationStatus = z.enum([
  "notStarted",
  "running",
  "succeeded",
  "failed",
  "unknownFutureValue"
]);
const microsoft_graph_publicErrorDetail = z.object({
  code: z.string().describe("The error code.").nullish(),
  message: z.string().describe("The error message.").nullish(),
  target: z.string().describe("The target of the error.").nullish()
}).strict();
const microsoft_graph_publicInnerError = z.object({
  code: z.string().describe("The error code.").nullish(),
  details: z.array(microsoft_graph_publicErrorDetail).describe("A collection of error details.").optional(),
  message: z.string().describe("The error message.").nullish(),
  target: z.string().describe("The target of the error.").nullish()
}).strict();
const microsoft_graph_publicError = z.object({
  code: z.string().describe("Represents the error code.").nullish(),
  details: z.array(microsoft_graph_publicErrorDetail).describe("Details of the error.").optional(),
  innerError: microsoft_graph_publicInnerError.optional(),
  message: z.string().describe("A non-localized message for the developer.").nullish(),
  target: z.string().describe("The target of the error.").nullish()
}).strict();
const microsoft_graph_richLongRunningOperation = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  createdDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The start time of the operation. The timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z."
  ).nullish(),
  lastActionDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "The time of the last action in the operation. The timestamp type represents date and time information using ISO 8601 format and is always in UTC. For example, midnight UTC on Jan 1, 2014 is 2014-01-01T00:00:00Z."
  ).nullish(),
  resourceLocation: z.string().describe("URI of the resource that the operation is performed on.").nullish(),
  status: microsoft_graph_longRunningOperationStatus.optional(),
  statusDetail: z.string().describe("Details about the status of the operation.").nullish(),
  error: microsoft_graph_publicError.optional(),
  percentageComplete: z.number().gte(-2147483648).lte(2147483647).describe("A value between 0 and 100 that indicates the progress of the operation.").nullish(),
  resourceId: z.string().describe("The unique identifier for the result.").nullish(),
  type: z.string().describe("The type of the operation.").nullish()
}).strict();
const microsoft_graph_subscription = z.object({
  id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
  applicationId: z.string().describe(
    "Optional. Identifier of the application used to create the subscription. Read-only."
  ).nullish(),
  changeType: z.string().describe(
    "Required. Indicates the type of change in the subscribed resource that raises a change notification. The supported values are: created, updated, deleted. Multiple values can be combined using a comma-separated list. Note:  Drive root item and list change notifications support only the updated changeType. User and group change notifications support updated and deleted changeType. Use updated to receive notifications when user or group is created, updated, or soft deleted. Use deleted to receive notifications when user or group is permanently deleted."
  ).optional(),
  clientState: z.string().describe(
    "Optional. Specifies the value of the clientState property sent by the service in each change notification. The maximum length is 128 characters. The client can check that the change notification came from the service by comparing the value of the clientState property sent with the subscription with the value of the clientState property received with each change notification."
  ).nullish(),
  creatorId: z.string().describe(
    "Optional. Identifier of the user or service principal that created the subscription. If the app used delegated permissions to create the subscription, this field contains the ID of the signed-in user the app called on behalf of. If the app used application permissions, this field contains the ID of the service principal corresponding to the app. Read-only."
  ).nullish(),
  encryptionCertificate: z.string().describe(
    "Optional. A base64-encoded representation of a certificate with a public key used to encrypt resource data in change notifications. Optional but required when includeResourceData is true."
  ).nullish(),
  encryptionCertificateId: z.string().describe(
    "Optional. A custom app-provided identifier to help identify the certificate needed to decrypt resource data."
  ).nullish(),
  expirationDateTime: z.string().regex(
    /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
  ).datetime({ offset: true }).describe(
    "Required. Specifies the date and time when the webhook subscription expires. The time is in UTC, and can be an amount of time from subscription creation that varies for the resource subscribed to. Any value under 45 minutes after the time of the request is automatically set to 45 minutes after the request time. For the maximum supported subscription length of time, see Subscription lifetime."
  ).optional(),
  includeResourceData: z.boolean().describe(
    "Optional. When set to true, change notifications include resource data (such as content of a chat message)."
  ).nullish(),
  latestSupportedTlsVersion: z.string().describe(
    "Optional. Specifies the latest version of Transport Layer Security (TLS) that the notification endpoint, specified by notificationUrl, supports. The possible values are: v10, v11, v12, v13. For subscribers whose notification endpoint supports a version lower than the currently recommended version (TLS 1.2), specifying this property by a set timeline allows them to temporarily use their deprecated version of TLS before completing their upgrade to TLS 1.2. For these subscribers, not setting this property per the timeline would result in subscription operations failing. For subscribers whose notification endpoint already supports TLS 1.2, setting this property is optional. In such cases, Microsoft Graph defaults the property to v1_2."
  ).nullish(),
  lifecycleNotificationUrl: z.string().describe(
    "Required for Teams resources if  the expirationDateTime value is more than 1 hour from now; optional otherwise. The URL of the endpoint that receives lifecycle notifications, including subscriptionRemoved, reauthorizationRequired, and missed notifications. This URL must make use of the HTTPS protocol. For more information, see Reduce missing subscriptions and change notifications."
  ).nullish(),
  notificationQueryOptions: z.string().describe(
    "Optional. OData query options for specifying value for the targeting resource. Clients receive notifications when resource reaches the state matching the query options provided here. With this new property in the subscription creation payload along with all existing properties, Webhooks deliver notifications whenever a resource reaches the desired state mentioned in the notificationQueryOptions property. For example, when the print job is completed or when a print job resource isFetchable property value becomes true etc.  Supported only for Universal Print Service. For more information, see Subscribe to change notifications from cloud printing APIs using Microsoft Graph."
  ).nullish(),
  notificationUrl: z.string().describe(
    "Required. The URL of the endpoint that receives the change notifications. This URL must make use of the HTTPS protocol. Any query string parameter included in the notificationUrl property is included in the HTTP POST request when Microsoft Graph sends the change notifications."
  ).optional(),
  notificationUrlAppId: z.string().describe(
    "Optional. The app ID that the subscription service can use to generate the validation token. The value allows the client to validate the authenticity of the notification received."
  ).nullish(),
  resource: z.string().describe(
    "Required. Specifies the resource that is monitored for changes. Don't include the base URL (https://graph.microsoft.com/v1.0/). See the possible resource path values for each supported resource."
  ).optional()
}).strict();
const microsoft_graph_list = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    createdBy: microsoft_graph_identitySet.optional(),
    createdDateTime: z.string().regex(
      /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
    ).datetime({ offset: true }).describe("Date and time of item creation. Read-only.").optional(),
    description: z.string().describe("Provides a user-visible description of the item. Optional.").nullish(),
    eTag: z.string().describe("ETag for the item. Read-only.").nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    lastModifiedDateTime: z.string().regex(
      /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
    ).datetime({ offset: true }).describe("Date and time the item was last modified. Read-only.").optional(),
    name: z.string().describe("The name of the item. Read-write.").nullish(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z.string().describe(
      "URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only."
    ).nullish(),
    createdByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
    lastModifiedByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
    displayName: z.string().describe("The displayable title of the list.").nullish(),
    list: microsoft_graph_listInfo.optional(),
    sharepointIds: microsoft_graph_sharepointIds.optional(),
    system: microsoft_graph_systemFacet.optional(),
    columns: z.array(microsoft_graph_columnDefinition).describe("The collection of field definitions for this list.").optional(),
    contentTypes: z.array(microsoft_graph_contentType).describe("The collection of content types present in this list.").optional(),
    drive: microsoft_graph_drive.optional(),
    items: z.array(microsoft_graph_listItem).describe("All items contained in the list.").optional(),
    operations: z.array(microsoft_graph_richLongRunningOperation).describe("The collection of long-running operations on the list.").optional(),
    subscriptions: z.array(microsoft_graph_subscription).describe("The set of subscriptions on the list.").optional()
  }).strict()
);
const microsoft_graph_drive = z.lazy(
  () => z.object({
    id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
    createdBy: microsoft_graph_identitySet.optional(),
    createdDateTime: z.string().regex(
      /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
    ).datetime({ offset: true }).describe("Date and time of item creation. Read-only.").optional(),
    description: z.string().describe("Provides a user-visible description of the item. Optional.").nullish(),
    eTag: z.string().describe("ETag for the item. Read-only.").nullish(),
    lastModifiedBy: microsoft_graph_identitySet.optional(),
    lastModifiedDateTime: z.string().regex(
      /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
    ).datetime({ offset: true }).describe("Date and time the item was last modified. Read-only.").optional(),
    name: z.string().describe("The name of the item. Read-write.").nullish(),
    parentReference: microsoft_graph_itemReference.optional(),
    webUrl: z.string().describe(
      "URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only."
    ).nullish(),
    createdByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
    lastModifiedByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
    driveType: z.string().describe(
      "Describes the type of drive represented by this resource. OneDrive personal drives return personal. OneDrive for Business returns business. SharePoint document libraries return documentLibrary. Read-only."
    ).nullish(),
    owner: microsoft_graph_identitySet.optional(),
    quota: microsoft_graph_quota.optional(),
    sharePointIds: microsoft_graph_sharepointIds.optional(),
    system: microsoft_graph_systemFacet.optional(),
    bundles: z.array(microsoft_graph_driveItem).describe(
      "Collection of bundles (albums and multi-select-shared sets of items). Only in personal OneDrive."
    ).optional(),
    following: z.array(microsoft_graph_driveItem).describe("The list of items the user is following. Only in OneDrive for Business.").optional(),
    items: z.array(microsoft_graph_driveItem).describe("All items contained in the drive. Read-only. Nullable.").optional(),
    list: microsoft_graph_list.optional(),
    root: microsoft_graph_driveItem.describe("[Note: Simplified from 45 properties to 25 most common ones]").optional(),
    special: z.array(microsoft_graph_driveItem).describe("Collection of common folders available in OneDrive. Read-only. Nullable.").optional()
  }).strict()
);
const microsoft_graph_driveCollectionResponse = z.object({
  "@odata.count": z.number().int().nullable(),
  "@odata.nextLink": z.string().nullable(),
  value: z.array(microsoft_graph_drive)
}).partial().strict();
const schemas = {
  microsoft_graph_ODataErrors_ErrorDetails,
  microsoft_graph_ODataErrors_InnerError,
  microsoft_graph_ODataErrors_MainError,
  microsoft_graph_ODataErrors_ODataError,
  microsoft_graph_identity,
  microsoft_graph_identitySet,
  microsoft_graph_sharepointIds,
  microsoft_graph_itemReference,
  microsoft_graph_assignedLicense,
  microsoft_graph_assignedPlan,
  microsoft_graph_authorizationInfo,
  microsoft_graph_customSecurityAttributeValue,
  microsoft_graph_user,
  microsoft_graph_audio,
  microsoft_graph_album,
  microsoft_graph_bundle,
  microsoft_graph_deleted,
  microsoft_graph_hashes,
  microsoft_graph_file,
  microsoft_graph_fileSystemInfo,
  microsoft_graph_folderView,
  microsoft_graph_folder,
  microsoft_graph_image,
  microsoft_graph_geoCoordinates,
  microsoft_graph_malware,
  microsoft_graph_package,
  microsoft_graph_pendingContentUpdate,
  microsoft_graph_pendingOperations,
  microsoft_graph_driveItem,
  microsoft_graph_driveItemCollectionResponse,
  microsoft_graph_workbookChartFont,
  microsoft_graph_workbookChartLineFormat,
  microsoft_graph_workbookChartAxisFormat,
  microsoft_graph_workbookChartGridlinesFormat,
  microsoft_graph_workbookChartGridlines,
  microsoft_graph_workbookChartAxisTitleFormat,
  microsoft_graph_workbookChartAxisTitle,
  microsoft_graph_workbookChartAxis,
  microsoft_graph_workbookChartAxes,
  microsoft_graph_workbookChartFill,
  microsoft_graph_workbookChartDataLabelFormat,
  microsoft_graph_workbookChartDataLabels,
  microsoft_graph_workbookChartAreaFormat,
  microsoft_graph_workbookChartLegendFormat,
  microsoft_graph_workbookChartLegend,
  microsoft_graph_workbookChartSeriesFormat,
  microsoft_graph_workbookChartPointFormat,
  microsoft_graph_workbookChartPoint,
  microsoft_graph_workbookChartSeries,
  microsoft_graph_workbookChartTitleFormat,
  microsoft_graph_workbookChartTitle,
  microsoft_graph_workbookChart,
  microsoft_graph_workbookNamedItem,
  microsoft_graph_workbookPivotTable,
  microsoft_graph_workbookWorksheetProtectionOptions,
  microsoft_graph_workbookWorksheetProtection,
  microsoft_graph_workbookIcon,
  microsoft_graph_workbookFilterCriteria,
  microsoft_graph_workbookFilter,
  microsoft_graph_workbookTableColumn,
  microsoft_graph_workbookTableRow,
  microsoft_graph_workbookSortField,
  microsoft_graph_workbookTableSort,
  microsoft_graph_workbookTable,
  microsoft_graph_workbookWorksheet,
  microsoft_graph_workbookWorksheetCollectionResponse,
  create_excel_chart_Body,
  microsoft_graph_workbookRangeBorder,
  microsoft_graph_workbookRangeFill,
  microsoft_graph_workbookRangeFont,
  microsoft_graph_workbookFormatProtection,
  microsoft_graph_workbookRangeFormat,
  microsoft_graph_workbookRangeSort,
  microsoft_graph_workbookRange,
  microsoft_graph_storagePlanInformation,
  microsoft_graph_quota,
  microsoft_graph_systemFacet,
  microsoft_graph_listInfo,
  microsoft_graph_columnTypes,
  microsoft_graph_booleanColumn,
  microsoft_graph_calculatedColumn,
  microsoft_graph_choiceColumn,
  microsoft_graph_contentApprovalStatusColumn,
  microsoft_graph_currencyColumn,
  microsoft_graph_dateTimeColumn,
  microsoft_graph_defaultColumnValue,
  microsoft_graph_geolocationColumn,
  microsoft_graph_hyperlinkOrPictureColumn,
  microsoft_graph_lookupColumn,
  microsoft_graph_numberColumn,
  microsoft_graph_personOrGroupColumn,
  microsoft_graph_columnDefinition,
  microsoft_graph_contentTypeInfo,
  microsoft_graph_documentSetContent,
  microsoft_graph_documentSet,
  microsoft_graph_contentTypeOrder,
  microsoft_graph_columnLink,
  microsoft_graph_contentType,
  microsoft_graph_itemActionStat,
  microsoft_graph_incompleteData,
  microsoft_graph_accessAction,
  microsoft_graph_itemActivity,
  microsoft_graph_itemActivityStat,
  microsoft_graph_itemAnalytics,
  microsoft_graph_publicationFacet,
  microsoft_graph_fieldValueSet,
  microsoft_graph_documentSetVersionItem,
  microsoft_graph_documentSetVersion,
  microsoft_graph_listItemVersion,
  microsoft_graph_listItem,
  microsoft_graph_longRunningOperationStatus,
  microsoft_graph_publicErrorDetail,
  microsoft_graph_publicInnerError,
  microsoft_graph_publicError,
  microsoft_graph_richLongRunningOperation,
  microsoft_graph_subscription,
  microsoft_graph_list,
  microsoft_graph_drive,
  microsoft_graph_driveCollectionResponse
};
const endpoints = makeApi([
  {
    method: "delete",
    path: "/drives/:driveId/items/:driveItemId",
    alias: "delete-onedrive-file",
    description: `Delete navigation property items for drives`,
    requestFormat: "json",
    parameters: [
      {
        name: "If-Match",
        type: "Header",
        schema: z.string().describe("ETag").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/drives/:driveId/items/:driveItemId/children",
    alias: "list-folder-files",
    description: `Return a collection of DriveItems in the children relationship of a DriveItem. DriveItems with a non-null folder or package facet can have one or more child DriveItems.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$top",
        type: "Query",
        schema: z.number().int().gte(0).describe("Show only the first n items").optional()
      },
      {
        name: "$skip",
        type: "Query",
        schema: z.number().int().gte(0).describe("Skip the first n items").optional()
      },
      {
        name: "$search",
        type: "Query",
        schema: z.string().describe("Search items by search phrases").optional()
      },
      {
        name: "$filter",
        type: "Query",
        schema: z.string().describe("Filter items by property values").optional()
      },
      {
        name: "$count",
        type: "Query",
        schema: z.boolean().describe("Include count of items").optional()
      },
      {
        name: "$orderby",
        type: "Query",
        schema: z.array(z.string()).describe("Order items by property values").optional()
      },
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "post",
    path: "/drives/:driveId/items/:driveItemId/children",
    alias: "create-folder",
    description: `Create new navigation property to children for drives`,
    requestFormat: "json",
    parameters: [
      {
        name: "body",
        description: `New navigation property`,
        type: "Body",
        schema: z.object({
          id: z.string().describe("The unique identifier for an entity. Read-only.").optional(),
          name: z.string().describe("The name of the item. Read-write.").nullish(),
          description: z.string().describe("Provides a user-visible description of the item. Optional.").nullish(),
          createdDateTime: z.string().regex(
            /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
          ).datetime({ offset: true }).describe("Date and time of item creation. Read-only.").optional(),
          lastModifiedDateTime: z.string().regex(
            /^[0-9]{4,}-(0[1-9]|1[012])-(0[1-9]|[12][0-9]|3[01])T([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]([.][0-9]{1,12})?(Z|[+-][0-9][0-9]:[0-9][0-9])$/
          ).datetime({ offset: true }).describe("Date and time the item was last modified. Read-only.").optional(),
          content: z.string().describe("The content stream, if the item represents a file.").nullish(),
          createdBy: microsoft_graph_identitySet.optional(),
          eTag: z.string().describe("ETag for the item. Read-only.").nullish(),
          lastModifiedBy: microsoft_graph_identitySet.optional(),
          parentReference: microsoft_graph_itemReference.optional(),
          webUrl: z.string().describe(
            "URL that either displays the resource in the browser (for Office file formats), or is a direct link to the file (for other formats). Read-only."
          ).nullish(),
          createdByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
          lastModifiedByUser: microsoft_graph_user.describe("[Note: Simplified from 135 properties to 25 most common ones]").optional(),
          audio: microsoft_graph_audio.optional(),
          bundle: microsoft_graph_bundle.optional(),
          cTag: z.string().describe(
            "An eTag for the content of the item. This eTag isn't changed if only the metadata is changed. Note This property isn't returned if the item is a folder. Read-only."
          ).nullish(),
          deleted: microsoft_graph_deleted.optional(),
          file: microsoft_graph_file.optional(),
          fileSystemInfo: microsoft_graph_fileSystemInfo.optional(),
          folder: microsoft_graph_folder.optional(),
          image: microsoft_graph_image.optional(),
          location: microsoft_graph_geoCoordinates.optional(),
          malware: microsoft_graph_malware.optional(),
          package: microsoft_graph_package.optional(),
          pendingOperations: microsoft_graph_pendingOperations.optional()
        }).strict().passthrough()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/drives/:driveId/items/:driveItemId/content",
    alias: "download-onedrive-file-content",
    description: `The content stream, if the item represents a file.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$format",
        type: "Query",
        schema: z.string().describe("Format of the content").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "put",
    path: "/drives/:driveId/items/:driveItemId/content",
    alias: "upload-file-content",
    description: `The content stream, if the item represents a file.`,
    requestFormat: "binary",
    parameters: [
      {
        name: "body",
        description: `New media content.`,
        type: "Body",
        schema: z.union([z.string(), z.instanceof(Buffer)]).optional()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/drives/:driveId/items/:driveItemId/workbook/worksheets",
    alias: "list-excel-worksheets",
    description: `Represents a collection of worksheets associated with the workbook. Read-only.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$top",
        type: "Query",
        schema: z.number().int().gte(0).describe("Show only the first n items").optional()
      },
      {
        name: "$skip",
        type: "Query",
        schema: z.number().int().gte(0).describe("Skip the first n items").optional()
      },
      {
        name: "$search",
        type: "Query",
        schema: z.string().describe("Search items by search phrases").optional()
      },
      {
        name: "$filter",
        type: "Query",
        schema: z.string().describe("Filter items by property values").optional()
      },
      {
        name: "$count",
        type: "Query",
        schema: z.boolean().describe("Include count of items").optional()
      },
      {
        name: "$orderby",
        type: "Query",
        schema: z.array(z.string()).describe("Order items by property values").optional()
      },
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "post",
    path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/charts/add",
    alias: "create-excel-chart",
    description: `Creates a new chart.`,
    requestFormat: "json",
    parameters: [
      {
        name: "body",
        description: `Action parameters`,
        type: "Body",
        schema: create_excel_chart_Body
      }
    ],
    response: z.void()
  },
  {
    method: "patch",
    path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range()/format",
    alias: "format-excel-range",
    description: `Update the navigation property format in drives`,
    requestFormat: "json",
    parameters: [
      {
        name: "body",
        description: `New navigation property values`,
        type: "Body",
        schema: microsoft_graph_workbookRangeFormat
      }
    ],
    response: z.void()
  },
  {
    method: "patch",
    path: "/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range()/sort",
    alias: "sort-excel-range",
    description: `Update the navigation property sort in drives`,
    requestFormat: "json",
    parameters: [
      {
        name: "body",
        description: `New navigation property values`,
        type: "Body",
        schema: z.object({
          id: z.string().describe("The unique identifier for an entity. Read-only.").optional()
        }).strict()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: `/drives/:driveId/items/:driveItemId/workbook/worksheets/:workbookWorksheetId/range(address=':address')`,
    alias: "get-excel-range",
    description: `Invoke function range`,
    requestFormat: "json",
    response: z.void()
  },
  {
    method: "get",
    path: "/drives/:driveId/root",
    alias: "get-drive-root-item",
    description: `The root folder of the drive. Read-only.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/me/drives",
    alias: "list-drives",
    description: `Retrieve the list of Drive resources available for a target User, Group, or Site.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$top",
        type: "Query",
        schema: z.number().int().gte(0).describe("Show only the first n items").optional()
      },
      {
        name: "$skip",
        type: "Query",
        schema: z.number().int().gte(0).describe("Skip the first n items").optional()
      },
      {
        name: "$search",
        type: "Query",
        schema: z.string().describe("Search items by search phrases").optional()
      },
      {
        name: "$filter",
        type: "Query",
        schema: z.string().describe("Filter items by property values").optional()
      },
      {
        name: "$count",
        type: "Query",
        schema: z.boolean().describe("Include count of items").optional()
      },
      {
        name: "$orderby",
        type: "Query",
        schema: z.array(z.string()).describe("Order items by property values").optional()
      },
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/sites/:siteId/drives",
    alias: "list-sharepoint-site-drives",
    description: `The collection of drives (document libraries) under this site.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$top",
        type: "Query",
        schema: z.number().int().gte(0).describe("Show only the first n items").optional()
      },
      {
        name: "$skip",
        type: "Query",
        schema: z.number().int().gte(0).describe("Skip the first n items").optional()
      },
      {
        name: "$search",
        type: "Query",
        schema: z.string().describe("Search items by search phrases").optional()
      },
      {
        name: "$filter",
        type: "Query",
        schema: z.string().describe("Filter items by property values").optional()
      },
      {
        name: "$count",
        type: "Query",
        schema: z.boolean().describe("Include count of items").optional()
      },
      {
        name: "$orderby",
        type: "Query",
        schema: z.array(z.string()).describe("Order items by property values").optional()
      },
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  },
  {
    method: "get",
    path: "/sites/:siteId/drives/:driveId",
    alias: "get-sharepoint-site-drive-by-id",
    description: `The collection of drives (document libraries) under this site.`,
    requestFormat: "json",
    parameters: [
      {
        name: "$select",
        type: "Query",
        schema: z.array(z.string()).describe("Select properties to be returned").optional()
      },
      {
        name: "$expand",
        type: "Query",
        schema: z.array(z.string()).describe("Expand related entities").optional()
      }
    ],
    response: z.void()
  }
]);
const api = new Zodios(endpoints);
function createApiClient(baseUrl, options) {
  return new Zodios(baseUrl, endpoints, options);
}
export {
  api,
  createApiClient,
  schemas
};
