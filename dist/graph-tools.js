import logger from "./logger.js";
import { api } from "./generated/client.js";
import { z } from "zod";
import { readFileSync } from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { readFileForUpload, validateFilePath } from "./lib/file-utils.js";
import { TOOL_CATEGORIES } from "./tool-categories.js";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, "endpoints.json"), "utf8")
);
async function executeGraphTool(tool, config, graphClient, params) {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
  try {
    const parameterDefinitions = tool.parameters || [];
    let path2 = tool.path;
    const queryParams = {};
    const headers = {};
    let body = null;
    for (const [paramName, paramValue] of Object.entries(params)) {
      if ([
        "fetchAllPages",
        "includeHeaders",
        "excludeResponse",
        "timezone",
        "expandExtendedProperties"
      ].includes(paramName)) {
        continue;
      }
      const odataParams = [
        "filter",
        "select",
        "expand",
        "orderby",
        "skip",
        "top",
        "count",
        "search",
        "format"
      ];
      const normalizedParamName = paramName.startsWith("$") ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      const paramDef = parameterDefinitions.find(
        (p) => p.name === paramName || isOdataParam && p.name === normalizedParamName
      );
      if (paramDef) {
        switch (paramDef.type) {
          case "Path": {
            const shouldSkipEncoding = config?.skipEncoding?.includes(paramName) ?? false;
            const encodedValue = shouldSkipEncoding ? paramValue : encodeURIComponent(paramValue);
            path2 = path2.replace(`{${paramName}}`, encodedValue).replace(`:${paramName}`, encodedValue);
            break;
          }
          case "Query":
            if (paramValue !== "" && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;
          case "Body":
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;
          case "Header":
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === "body") {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      }
    }
    if (config?.supportsFileUpload) {
      const hasFilePath = "filePath" in params && params.filePath !== void 0 && params.filePath !== null;
      const hasBody = "body" in params && params.body !== void 0 && params.body !== null;
      if (hasFilePath && hasBody) {
        const errorMsg = 'Cannot use both "filePath" and "body" parameters. Use "filePath" for local file upload or "body" for direct content.';
        logger.error(errorMsg);
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ error: errorMsg })
            }
          ],
          isError: true
        };
      }
      if (hasFilePath) {
        logger.info(`Processing file upload from local path: ${params.filePath}`);
        try {
          const fileData = readFileForUpload(params.filePath);
          body = fileData.content;
          headers["Content-Type"] = fileData.mimeType;
          logger.info(
            `File upload prepared: ${fileData.fileName} (${fileData.size} bytes, ${fileData.mimeType})`
          );
          delete params.filePath;
        } catch (error) {
          logger.error(`File upload preparation failed: ${error.message}`);
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  error: `File upload failed: ${error.message}`
                })
              }
            ],
            isError: true
          };
        }
      }
    }
    if (config?.supportsTimezone && params.timezone) {
      headers["Prefer"] = `outlook.timezone="${params.timezone}"`;
      logger.info(`Setting timezone header: Prefer: outlook.timezone="${params.timezone}"`);
    }
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = "singleValueExtendedProperties";
      if (queryParams["$expand"]) {
        queryParams["$expand"] += `,${expandValue}`;
      } else {
        queryParams["$expand"] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }
    if (config?.contentType) {
      headers["Content-Type"] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }
    if (config?.defaultTop && !queryParams["$top"]) {
      queryParams["$top"] = `${config.defaultTop}`;
      logger.info(`Forcing pagination: injected $top=${config.defaultTop} for ${tool.alias}`);
    }
    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams).map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`).join("&");
      path2 = `${path2}${path2.includes("?") ? "&" : "?"}${queryString}`;
    }
    const options = {
      method: tool.method.toUpperCase(),
      headers
    };
    if (options.method !== "GET" && body) {
      if (Buffer.isBuffer(body)) {
        options.body = body;
        logger.info(`Sending binary content (${body.length} bytes)`);
      } else if (config?.contentType === "text/html") {
        if (typeof body === "string") {
          options.body = body;
        } else if (typeof body === "object" && "content" in body) {
          options.body = body.content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === "string" ? body : JSON.stringify(body);
      }
    }
    const isProbablyMediaContent = tool.errors?.some((error) => error.description === "Retrieved media content") || path2.endsWith("/content");
    if (config?.returnDownloadUrl && path2.endsWith("/content")) {
      path2 = path2.replace(/\/content$/, "");
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }
    logger.info(`Making graph request to ${path2} with options: ${JSON.stringify(options)}`);
    let response = await graphClient.graphRequest(path2, options);
    const fetchAllPages = params.fetchAllPages === true;
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse["@odata.nextLink"];
        let pageCount = 1;
        while (nextLink && pageCount < 100) {
          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);
          const url = new URL(nextLink);
          const nextPath = url.pathname.replace("/v1.0", "");
          const nextOptions = { ...options };
          const nextQueryParams = {};
          for (const [key, value] of url.searchParams.entries()) {
            nextQueryParams[key] = value;
          }
          nextOptions.queryParams = nextQueryParams;
          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse["@odata.nextLink"];
            pageCount++;
          } else {
            break;
          }
        }
        if (pageCount >= 100) {
          logger.warn(`Reached maximum page limit (100) for pagination`);
        }
        combinedResponse.value = allItems;
        if (combinedResponse["@odata.count"]) {
          combinedResponse["@odata.count"] = allItems.length;
        }
        delete combinedResponse["@odata.nextLink"];
        response.content[0].text = JSON.stringify(combinedResponse);
        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }
    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);
      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse["@odata.nextLink"]) {
          logger.info(`Response has pagination nextLink: ${jsonResponse["@odata.nextLink"]}`);
        }
      } catch {
      }
    }
    const content = response.content.map((item) => ({
      type: "text",
      text: item.text
    }));
    return {
      content,
      _meta: response._meta,
      isError: response.isError
    };
  } catch (error) {
    logger.error(`Error in tool ${tool.alias}: ${error.message}`);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${error.message}`
          })
        }
      ],
      isError: true
    };
  }
}
function registerGraphTools(server, graphClient, readOnly = false, enabledToolsPattern, orgMode = false) {
  let enabledToolsRegex;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, "i");
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }
  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }
    if (readOnly && tool.method.toUpperCase() !== "GET") {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      skippedCount++;
      continue;
    }
    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }
    const paramSchema = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }
    if (tool.method.toUpperCase() === "GET" && tool.path.includes("/")) {
      paramSchema["fetchAllPages"] = z.boolean().describe("Automatically fetch all pages of results").optional();
    }
    paramSchema["includeHeaders"] = z.boolean().describe("Include response headers (including ETag) in the response metadata").optional();
    paramSchema["excludeResponse"] = z.boolean().describe("Exclude the full response body and only return success or failure indication").optional();
    if (endpointConfig?.supportsTimezone) {
      paramSchema["timezone"] = z.string().describe(
        'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
      ).optional();
    }
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema["expandExtendedProperties"] = z.boolean().describe(
        "When true, expands singleValueExtendedProperties on each event. Use this to retrieve custom extended properties (e.g., sync metadata) stored on calendar events."
      ).optional();
    }
    if (endpointConfig?.supportsFileUpload) {
      paramSchema["filePath"] = z.string().refine(validateFilePath, {
        message: "filePath must be a non-empty string"
      }).describe(
        'Local file path to upload. File content and MIME type are automatically detected. Cannot be used with "body" parameter.'
      ).optional();
    }
    let toolDescription = tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `

\u{1F4A1} TIP: ${endpointConfig.llmTip}`;
    }
    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === "GET",
          destructiveHint: ["POST", "PATCH", "DELETE"].includes(tool.method.toUpperCase()),
          openWorldHint: true
          // All tools call Microsoft Graph API
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${error.message}`);
      failedCount++;
    }
  }
  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}
function buildToolsRegistry(readOnly, orgMode) {
  const toolsMap = /* @__PURE__ */ new Map();
  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }
    if (readOnly && tool.method.toUpperCase() !== "GET") {
      continue;
    }
    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }
  return toolsMap;
}
function registerDiscoveryTools(server, graphClient, readOnly = false, orgMode = false) {
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode);
  logger.info(`Discovery mode: ${toolsRegistry.size} tools available in registry`);
  server.tool(
    "search-tools",
    `Search through ${toolsRegistry.size} available Microsoft Graph API tools. Use this to find tools by name, path, or description before executing them.`,
    {
      query: z.string().describe("Search query to filter tools (searches name, path, and description)").optional(),
      category: z.string().describe(
        "Filter by category: mail, calendar, files, contacts, tasks, onenote, search, users, excel"
      ).optional(),
      limit: z.number().describe("Maximum results to return (default: 20, max: 50)").optional()
    },
    {
      title: "search-tools",
      readOnlyHint: true,
      openWorldHint: true
      // Searches Microsoft Graph API tools
    },
    async ({ query, category, limit = 20 }) => {
      const maxLimit = Math.min(limit, 50);
      const results = [];
      const queryLower = query?.toLowerCase();
      const categoryDef = category ? TOOL_CATEGORIES[category] : void 0;
      for (const [name, { tool, config }] of toolsRegistry) {
        if (categoryDef && !categoryDef.pattern.test(name)) {
          continue;
        }
        if (queryLower) {
          const searchText = `${name} ${tool.path} ${tool.description || ""} ${config?.llmTip || ""}`.toLowerCase();
          if (!searchText.includes(queryLower)) {
            continue;
          }
        }
        results.push({
          name,
          method: tool.method.toUpperCase(),
          path: tool.path,
          description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`
        });
        if (results.length >= maxLimit) break;
      }
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                found: results.length,
                total: toolsRegistry.size,
                tools: results,
                tip: "Use execute-tool with the tool name and required parameters to call any of these tools."
              },
              null,
              2
            )
          }
        ]
      };
    }
  );
  server.tool(
    "execute-tool",
    "Execute a Microsoft Graph API tool by name. Use search-tools first to find available tools and their parameters.",
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z.record(z.any()).describe("Parameters to pass to the tool as key-value pairs").optional()
    },
    {
      title: "execute-tool",
      readOnlyHint: false,
      destructiveHint: true,
      // Can execute any tool, including write operations
      openWorldHint: true
      // Executes against Microsoft Graph API
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (!toolData) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: "Use search-tools to find available tools."
              })
            }
          ],
          isError: true
        };
      }
      return executeGraphTool(toolData.tool, toolData.config, graphClient, parameters);
    }
  );
}
export {
  registerDiscoveryTools,
  registerGraphTools
};
