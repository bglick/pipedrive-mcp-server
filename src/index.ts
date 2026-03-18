import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import * as pipedrive from "pipedrive";
import * as dotenv from 'dotenv';
import Bottleneck from 'bottleneck';
import jwt from 'jsonwebtoken';
import http from 'http';

// Type for error handling
interface ErrorWithMessage {
  message: string;
}

function isErrorWithMessage(error: unknown): error is ErrorWithMessage {
  return (
    typeof error === 'object' &&
    error !== null &&
    'message' in error &&
    typeof (error as Record<string, unknown>).message === 'string'
  );
}

function getErrorMessage(error: unknown): string {
  if (isErrorWithMessage(error)) {
    return error.message;
  }
  return String(error);
}

// Load environment variables
dotenv.config();

// Check for required environment variables
if (!process.env.PIPEDRIVE_API_TOKEN) {
  console.error("ERROR: PIPEDRIVE_API_TOKEN environment variable is required");
  process.exit(1);
}

if (!process.env.PIPEDRIVE_DOMAIN) {
  console.error("ERROR: PIPEDRIVE_DOMAIN environment variable is required (e.g., 'ukkofi.pipedrive.com')");
  process.exit(1);
}

const jwtSecret = process.env.MCP_JWT_SECRET;
const jwtAlgorithm = (process.env.MCP_JWT_ALGORITHM || 'HS256') as jwt.Algorithm;
const jwtVerifyOptions = {
  algorithms: [jwtAlgorithm],
  audience: process.env.MCP_JWT_AUDIENCE,
  issuer: process.env.MCP_JWT_ISSUER,
};

if (jwtSecret) {
  const bootToken = process.env.MCP_JWT_TOKEN;
  if (!bootToken) {
    console.error("ERROR: MCP_JWT_TOKEN environment variable is required when MCP_JWT_SECRET is set");
    process.exit(1);
  }

  try {
    jwt.verify(bootToken, jwtSecret, jwtVerifyOptions);
  } catch (error) {
    console.error("ERROR: Failed to verify MCP_JWT_TOKEN", error);
    process.exit(1);
  }
}

const verifyRequestAuthentication = (req: http.IncomingMessage) => {
  if (!jwtSecret) {
    return { ok: true } as const;
  }

  const header = req.headers['authorization'];
  if (!header) {
    return { ok: false, status: 401, message: 'Missing Authorization header' } as const;
  }

  const [scheme, token] = header.split(' ');
  if (scheme !== 'Bearer' || !token) {
    return { ok: false, status: 401, message: 'Invalid Authorization header format' } as const;
  }

  try {
    jwt.verify(token, jwtSecret, jwtVerifyOptions);
    return { ok: true } as const;
  } catch (error) {
    return { ok: false, status: 401, message: 'Invalid or expired token' } as const;
  }
};

const limiter = new Bottleneck({
  minTime: Number(process.env.PIPEDRIVE_RATE_LIMIT_MIN_TIME_MS || 250),
  maxConcurrent: Number(process.env.PIPEDRIVE_RATE_LIMIT_MAX_CONCURRENT || 2),
});

const withRateLimit = <T extends object>(client: T): T => {
  return new Proxy(client, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, receiver);
      if (typeof value === 'function') {
        return (...args: unknown[]) => limiter.schedule(() => (value as Function).apply(target, args));
      }
      return value;
    },
  });
};

// Initialize Pipedrive API client with API token and custom domain
const apiClient = new pipedrive.ApiClient();
apiClient.basePath = `https://${process.env.PIPEDRIVE_DOMAIN}/api/v1`;
apiClient.authentications = apiClient.authentications || {};
apiClient.authentications['api_key'] = {
  type: 'apiKey',
  'in': 'query',
  name: 'api_token',
  apiKey: process.env.PIPEDRIVE_API_TOKEN
};

// Initialize Pipedrive API clients
const dealsApi = withRateLimit(new pipedrive.DealsApi(apiClient));
const personsApi = withRateLimit(new pipedrive.PersonsApi(apiClient));
const organizationsApi = withRateLimit(new pipedrive.OrganizationsApi(apiClient));
const pipelinesApi = withRateLimit(new pipedrive.PipelinesApi(apiClient));
const itemSearchApi = withRateLimit(new pipedrive.ItemSearchApi(apiClient));
const leadsApi = withRateLimit(new pipedrive.LeadsApi(apiClient));
// @ts-ignore - ActivitiesApi exists but may not be in type definitions
const activitiesApi = withRateLimit(new pipedrive.ActivitiesApi(apiClient));
// @ts-ignore - NotesApi exists but may not be in type definitions
const notesApi = withRateLimit(new pipedrive.NotesApi(apiClient));
// @ts-ignore - UsersApi exists but may not be in type definitions
const usersApi = withRateLimit(new pipedrive.UsersApi(apiClient));

// Create MCP server
const server = new McpServer({
  name: "pipedrive-mcp-server",
  version: "1.0.2",
  capabilities: {
    resources: {},
    tools: {},
    prompts: {}
  }
});

// === TOOLS ===

// Get all users (for finding owner IDs)
server.tool(
  "get-users",
  "Get all users/owners from Pipedrive to identify owner IDs for filtering deals",
  {},
  async () => {
    try {
      const response = await usersApi.getUsers();
      const users = response.data?.map((user: any) => ({
        id: user.id,
        name: user.name,
        email: user.email,
        active_flag: user.active_flag,
        role_name: user.role_name
      })) || [];

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            summary: `Found ${users.length} users in your Pipedrive account`,
            users: users
          }, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching users:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching users: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get deals with flexible filtering options
server.tool(
  "get-deals",
  "Get deals from Pipedrive with flexible filtering options including search by title, date range, owner, stage, status, and more. Use 'get-users' tool first to find owner IDs.",
  {
    searchTitle: z.string().optional().describe("Search deals by title/name (partial matches supported)"),
    daysBack: z.number().optional().describe("Number of days back to fetch deals based on last activity date (default: 365)"),
    ownerId: z.number().optional().describe("Filter deals by owner/user ID (use get-users tool to find IDs)"),
    stageId: z.number().optional().describe("Filter deals by stage ID"),
    status: z.enum(['open', 'won', 'lost', 'deleted']).optional().describe("Filter deals by status (default: open)"),
    pipelineId: z.number().optional().describe("Filter deals by pipeline ID"),
    minValue: z.number().optional().describe("Minimum deal value filter"),
    maxValue: z.number().optional().describe("Maximum deal value filter"),
    limit: z.number().optional().describe("Maximum number of deals to return (default: 500)")
  },
  async ({
    searchTitle,
    daysBack = 365,
    ownerId,
    stageId,
    status = 'open',
    pipelineId,
    minValue,
    maxValue,
    limit = 500
  }) => {
    try {
      let filteredDeals: any[] = [];

      // If searching by title, use the search API first
      if (searchTitle) {
        // @ts-ignore - Bypass incorrect TypeScript definition
        const searchResponse = await dealsApi.searchDeals(searchTitle);
        filteredDeals = searchResponse.data || [];
      } else {
        // Calculate the date filter (daysBack days ago)
        const filterDate = new Date();
        filterDate.setDate(filterDate.getDate() - daysBack);
        const startDate = filterDate.toISOString().split('T')[0]; // Format as YYYY-MM-DD

        // Build API parameters (using actual Pipedrive API parameter names)
        const params: any = {
          sort: 'last_activity_date DESC',
          status: status,
          limit: limit
        };

        // Add optional filters
        if (ownerId) params.user_id = ownerId;
        if (stageId) params.stage_id = stageId;
        if (pipelineId) params.pipeline_id = pipelineId;

        // Fetch deals with filters
        // @ts-ignore - getDeals accepts parameters but types may be incomplete
        const response = await dealsApi.getDeals(params);
        filteredDeals = response.data || [];
      }

      // Apply additional client-side filtering

      // Filter by date if not searching by title
      if (!searchTitle) {
        const filterDate = new Date();
        filterDate.setDate(filterDate.getDate() - daysBack);

        filteredDeals = filteredDeals.filter((deal: any) => {
          if (!deal.last_activity_date) return false;
          const dealActivityDate = new Date(deal.last_activity_date);
          return dealActivityDate >= filterDate;
        });
      }

      // Filter by owner if specified and not already applied in API call
      if (ownerId && searchTitle) {
        filteredDeals = filteredDeals.filter((deal: any) => deal.owner_id === ownerId);
      }

      // Filter by status if specified and searching by title
      if (status && searchTitle) {
        filteredDeals = filteredDeals.filter((deal: any) => deal.status === status);
      }

      // Filter by stage if specified and not already applied in API call
      if (stageId && (searchTitle || !stageId)) {
        filteredDeals = filteredDeals.filter((deal: any) => deal.stage_id === stageId);
      }

      // Filter by pipeline if specified and not already applied in API call
      if (pipelineId && (searchTitle || !pipelineId)) {
        filteredDeals = filteredDeals.filter((deal: any) => deal.pipeline_id === pipelineId);
      }

      // Filter by value range if specified
      if (minValue !== undefined || maxValue !== undefined) {
        filteredDeals = filteredDeals.filter((deal: any) => {
          const value = parseFloat(deal.value) || 0;
          if (minValue !== undefined && value < minValue) return false;
          if (maxValue !== undefined && value > maxValue) return false;
          return true;
        });
      }

      // Apply limit
      if (filteredDeals.length > limit) {
        filteredDeals = filteredDeals.slice(0, limit);
      }

      // Build filter summary for response
      const filterSummary = {
        ...(searchTitle && { search_title: searchTitle }),
        ...(!searchTitle && { days_back: daysBack }),
        ...(!searchTitle && { filter_date: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000).toISOString().split('T')[0] }),
        status: status,
        ...(ownerId && { owner_id: ownerId }),
        ...(stageId && { stage_id: stageId }),
        ...(pipelineId && { pipeline_id: pipelineId }),
        ...(minValue !== undefined && { min_value: minValue }),
        ...(maxValue !== undefined && { max_value: maxValue }),
        total_deals_found: filteredDeals.length,
        limit_applied: limit
      };

      // Summarize deals to avoid massive responses but include notes and booking details
      const bookingFieldKey = "8f4b27fbd9dfc70d2296f23ce76987051ad7324e";
      const summarizedDeals = filteredDeals.map((deal: any) => ({
        id: deal.id,
        title: deal.title,
        value: deal.value,
        currency: deal.currency,
        status: deal.status,
        stage_name: deal.stage?.name || 'Unknown',
        pipeline_name: deal.pipeline?.name || 'Unknown',
        owner_name: deal.owner?.name || 'Unknown',
        organization_name: deal.org?.name || null,
        person_name: deal.person?.name || null,
        add_time: deal.add_time,
        last_activity_date: deal.last_activity_date,
        close_time: deal.close_time,
        won_time: deal.won_time,
        lost_time: deal.lost_time,
        notes_count: deal.notes_count || 0,
        // Include recent notes if available
        notes: deal.notes || [],
        // Include custom booking details field
        booking_details: deal[bookingFieldKey] || null
      }));

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            summary: searchTitle
              ? `Found ${filteredDeals.length} deals matching title search "${searchTitle}"`
              : `Found ${filteredDeals.length} deals matching the specified filters`,
            filters_applied: filterSummary,
            total_found: filteredDeals.length,
            deals: summarizedDeals.slice(0, 30) // Limit to 30 deals max to prevent huge responses
          }, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching deals:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching deals: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get deal by ID
server.tool(
  "get-deal",
  "Get a specific deal by ID including custom fields",
  {
    dealId: z.number().describe("Pipedrive deal ID")
  },
  async ({ dealId }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition, API expects just the ID
      const response = await dealsApi.getDeal(dealId);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error fetching deal ${dealId}:`, error);
      return {
        content: [{
          type: "text",
          text: `Error fetching deal ${dealId}: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get deal notes and custom booking details
server.tool(
  "get-deal-notes",
  "Get detailed notes and custom booking details for a specific deal",
  {
    dealId: z.number().describe("Pipedrive deal ID"),
    limit: z.number().optional().describe("Maximum number of notes to return (default: 20)")
  },
  async ({ dealId, limit = 20 }) => {
    try {
      const result: any = {
        deal_id: dealId,
        notes: [],
        booking_details: null
      };

      // Get deal details including custom fields
      try {
        // @ts-ignore - Bypass incorrect TypeScript definition
        const dealResponse = await dealsApi.getDeal(dealId);
        const deal = dealResponse.data;

        // Extract custom booking field
        const bookingFieldKey = "8f4b27fbd9dfc70d2296f23ce76987051ad7324e";
        if (deal && deal[bookingFieldKey]) {
          result.booking_details = deal[bookingFieldKey];
        }
      } catch (dealError) {
        console.error(`Error fetching deal details for ${dealId}:`, dealError);
        result.deal_error = getErrorMessage(dealError);
      }

      // Get deal notes
      try {
        // @ts-ignore - API parameters may not be fully typed
        // @ts-ignore - Bypass incorrect TypeScript definition
        const notesResponse = await notesApi.getNotes({
          deal_id: dealId,
          limit: limit
        });
        result.notes = notesResponse.data || [];
      } catch (noteError) {
        console.error(`Error fetching notes for deal ${dealId}:`, noteError);
        result.notes_error = getErrorMessage(noteError);
      }

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            summary: `Retrieved ${result.notes.length} notes and booking details for deal ${dealId}`,
            ...result
          }, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error fetching deal notes ${dealId}:`, error);
      return {
        content: [{
          type: "text",
          text: `Error fetching deal notes ${dealId}: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Search deals
server.tool(
  "search-deals",
  "Search deals by term",
  {
    term: z.string().describe("Search term for deals")
  },
  async ({ term }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await dealsApi.searchDeals(term);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error searching deals with term "${term}":`, error);
      return {
        content: [{
          type: "text",
          text: `Error searching deals: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get all persons
server.tool(
  "get-persons",
  "Get all persons from Pipedrive including custom fields",
  {},
  async () => {
    try {
      const response = await personsApi.getPersons();
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching persons:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching persons: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get person by ID
server.tool(
  "get-person",
  "Get a specific person by ID including custom fields",
  {
    personId: z.number().describe("Pipedrive person ID")
  },
  async ({ personId }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await personsApi.getPerson(personId);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error fetching person ${personId}:`, error);
      return {
        content: [{
          type: "text",
          text: `Error fetching person ${personId}: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Search persons
server.tool(
  "search-persons",
  "Search persons by term",
  {
    term: z.string().describe("Search term for persons")
  },
  async ({ term }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await personsApi.searchPersons(term);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error searching persons with term "${term}":`, error);
      return {
        content: [{
          type: "text",
          text: `Error searching persons: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get all organizations
server.tool(
  "get-organizations",
  "Get all organizations from Pipedrive including custom fields",
  {},
  async () => {
    try {
      const response = await organizationsApi.getOrganizations();
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching organizations:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching organizations: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get organization by ID
server.tool(
  "get-organization",
  "Get a specific organization by ID including custom fields",
  {
    organizationId: z.number().describe("Pipedrive organization ID")
  },
  async ({ organizationId }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await organizationsApi.getOrganization(organizationId);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error fetching organization ${organizationId}:`, error);
      return {
        content: [{
          type: "text",
          text: `Error fetching organization ${organizationId}: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Search organizations
server.tool(
  "search-organizations",
  "Search organizations by term",
  {
    term: z.string().describe("Search term for organizations")
  },
  async ({ term }) => {
    try {
      // @ts-ignore - API method exists but TypeScript definition is wrong
      const response = await (organizationsApi as any).searchOrganization({ term });
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error searching organizations with term "${term}":`, error);
      return {
        content: [{
          type: "text",
          text: `Error searching organizations: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get all pipelines
server.tool(
  "get-pipelines",
  "Get all pipelines from Pipedrive",
  {},
  async () => {
    try {
      const response = await pipelinesApi.getPipelines();
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching pipelines:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching pipelines: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get pipeline by ID
server.tool(
  "get-pipeline",
  "Get a specific pipeline by ID",
  {
    pipelineId: z.number().describe("Pipedrive pipeline ID")
  },
  async ({ pipelineId }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await pipelinesApi.getPipeline(pipelineId);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error fetching pipeline ${pipelineId}:`, error);
      return {
        content: [{
          type: "text",
          text: `Error fetching pipeline ${pipelineId}: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Get all stages
server.tool(
  "get-stages",
  "Get all stages from Pipedrive",
  {},
  async () => {
    try {
      // Since the stages are related to pipelines, we'll get all pipelines first
      const pipelinesResponse = await pipelinesApi.getPipelines();
      const pipelines = pipelinesResponse.data || [];
      
      // For each pipeline, fetch its stages
      const allStages = [];
      for (const pipeline of pipelines) {
        try {
          // @ts-ignore - Type definitions for getPipelineStages are incomplete
          const stagesResponse = await pipelinesApi.getPipelineStages(pipeline.id);
          const stagesData = Array.isArray(stagesResponse?.data)
            ? stagesResponse.data
            : [];

          if (stagesData.length > 0) {
            const pipelineStages = stagesData.map((stage: any) => ({
              ...stage,
              pipeline_name: pipeline.name
            }));
            allStages.push(...pipelineStages);
          }
        } catch (e) {
          console.error(`Error fetching stages for pipeline ${pipeline.id}:`, e);
        }
      }
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify(allStages, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error fetching stages:", error);
      return {
        content: [{
          type: "text",
          text: `Error fetching stages: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Search leads
server.tool(
  "search-leads",
  "Search leads by term",
  {
    term: z.string().describe("Search term for leads")
  },
  async ({ term }) => {
    try {
      // @ts-ignore - Bypass incorrect TypeScript definition
      const response = await leadsApi.searchLeads(term);
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error searching leads with term "${term}":`, error);
      return {
        content: [{
          type: "text",
          text: `Error searching leads: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// Generic search across item types
server.tool(
  "search-all",
  "Search across all item types (deals, persons, organizations, etc.)",
  {
    term: z.string().describe("Search term"),
    itemTypes: z.string().optional().describe("Comma-separated list of item types to search (deal,person,organization,product,file,activity,lead)")
  },
  async ({ term, itemTypes }) => {
    try {
      const itemType = itemTypes; // Just rename the parameter
      const response = await itemSearchApi.searchItem({ 
        term,
        itemType 
      });
      return {
        content: [{
          type: "text",
          text: JSON.stringify(response.data, null, 2)
        }]
      };
    } catch (error) {
      console.error(`Error performing search with term "${term}":`, error);
      return {
        content: [{
          type: "text",
          text: `Error performing search: ${getErrorMessage(error)}`
        }],
        isError: true
      };
    }
  }
);

// === WRITE TOOLS ===

// Add a deal
server.tool(
  "add-deal",
  "Create a new deal in Pipedrive",
  {
    title: z.string().describe("Deal title"),
    value: z.number().optional().describe("Deal value"),
    currency: z.string().optional().describe("Currency code (e.g. USD, EUR)"),
    status: z.enum(['open', 'won', 'lost']).optional().describe("Deal status (default: open)"),
    person_id: z.number().optional().describe("ID of the person linked to the deal"),
    org_id: z.number().optional().describe("ID of the organization linked to the deal"),
    pipeline_id: z.number().optional().describe("Pipeline ID"),
    stage_id: z.number().optional().describe("Stage ID"),
    expected_close_date: z.string().optional().describe("Expected close date (YYYY-MM-DD)"),
    lost_reason: z.string().optional().describe("Reason for losing the deal"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await dealsApi.addDeal(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding deal:", error);
      return { content: [{ type: "text", text: `Error adding deal: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update a deal
server.tool(
  "update-deal",
  "Update an existing deal in Pipedrive",
  {
    dealId: z.number().describe("Pipedrive deal ID"),
    title: z.string().optional().describe("Deal title"),
    value: z.number().optional().describe("Deal value"),
    currency: z.string().optional().describe("Currency code (e.g. USD, EUR)"),
    status: z.enum(['open', 'won', 'lost']).optional().describe("Deal status"),
    person_id: z.number().optional().describe("ID of the person linked to the deal"),
    org_id: z.number().optional().describe("ID of the organization linked to the deal"),
    pipeline_id: z.number().optional().describe("Pipeline ID"),
    stage_id: z.number().optional().describe("Stage ID"),
    expected_close_date: z.string().optional().describe("Expected close date (YYYY-MM-DD)"),
    lost_reason: z.string().optional().describe("Reason for losing the deal"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async ({ dealId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await dealsApi.updateDeal(dealId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating deal ${dealId}:`, error);
      return { content: [{ type: "text", text: `Error updating deal ${dealId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Add a person
server.tool(
  "add-person",
  "Create a new person in Pipedrive",
  {
    name: z.string().describe("Person's full name"),
    email: z.array(z.string()).optional().describe("Email address(es)"),
    phone: z.array(z.string()).optional().describe("Phone number(s)"),
    org_id: z.number().optional().describe("Organization ID to link this person to"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await personsApi.addPerson(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding person:", error);
      return { content: [{ type: "text", text: `Error adding person: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update a person
server.tool(
  "update-person",
  "Update an existing person in Pipedrive",
  {
    personId: z.number().describe("Pipedrive person ID"),
    name: z.string().optional().describe("Person's full name"),
    email: z.array(z.string()).optional().describe("Email address(es)"),
    phone: z.array(z.string()).optional().describe("Phone number(s)"),
    org_id: z.number().optional().describe("Organization ID"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async ({ personId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await personsApi.updatePerson(personId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating person ${personId}:`, error);
      return { content: [{ type: "text", text: `Error updating person ${personId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Add an organization
server.tool(
  "add-organization",
  "Create a new organization in Pipedrive",
  {
    name: z.string().describe("Organization name"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await organizationsApi.addOrganization(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding organization:", error);
      return { content: [{ type: "text", text: `Error adding organization: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update an organization
server.tool(
  "update-organization",
  "Update an existing organization in Pipedrive",
  {
    organizationId: z.number().describe("Pipedrive organization ID"),
    name: z.string().optional().describe("Organization name"),
    visible_to: z.number().optional().describe("Visibility (1=owner, 3=entire company)")
  },
  async ({ organizationId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await organizationsApi.updateOrganization(organizationId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating organization ${organizationId}:`, error);
      return { content: [{ type: "text", text: `Error updating organization ${organizationId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Add a note
server.tool(
  "add-note",
  "Add a note to a deal, person, or organization in Pipedrive",
  {
    content: z.string().describe("Note content (HTML supported)"),
    deal_id: z.number().optional().describe("Deal ID to attach the note to"),
    person_id: z.number().optional().describe("Person ID to attach the note to"),
    org_id: z.number().optional().describe("Organization ID to attach the note to"),
    pinned_to_deal_flag: z.boolean().optional().describe("Pin note to deal"),
    pinned_to_person_flag: z.boolean().optional().describe("Pin note to person"),
    pinned_to_organization_flag: z.boolean().optional().describe("Pin note to organization")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await notesApi.addNote(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding note:", error);
      return { content: [{ type: "text", text: `Error adding note: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update a note
server.tool(
  "update-note",
  "Update an existing note in Pipedrive",
  {
    noteId: z.number().describe("Pipedrive note ID"),
    content: z.string().optional().describe("Note content (HTML supported)"),
    pinned_to_deal_flag: z.boolean().optional().describe("Pin note to deal"),
    pinned_to_person_flag: z.boolean().optional().describe("Pin note to person"),
    pinned_to_organization_flag: z.boolean().optional().describe("Pin note to organization")
  },
  async ({ noteId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await notesApi.updateNote(noteId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating note ${noteId}:`, error);
      return { content: [{ type: "text", text: `Error updating note ${noteId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Add an activity
server.tool(
  "add-activity",
  "Create a new activity (call, meeting, task, etc.) in Pipedrive",
  {
    subject: z.string().describe("Activity subject/title"),
    type: z.string().optional().describe("Activity type (call, meeting, task, deadline, email, lunch)"),
    due_date: z.string().optional().describe("Due date (YYYY-MM-DD)"),
    due_time: z.string().optional().describe("Due time (HH:MM)"),
    duration: z.string().optional().describe("Duration (HH:MM)"),
    deal_id: z.number().optional().describe("Deal ID to link this activity to"),
    person_id: z.number().optional().describe("Person ID to link this activity to"),
    org_id: z.number().optional().describe("Organization ID to link this activity to"),
    note: z.string().optional().describe("Activity note"),
    done: z.boolean().optional().describe("Whether activity is done")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await activitiesApi.addActivity(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding activity:", error);
      return { content: [{ type: "text", text: `Error adding activity: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update an activity
server.tool(
  "update-activity",
  "Update an existing activity in Pipedrive",
  {
    activityId: z.number().describe("Pipedrive activity ID"),
    subject: z.string().optional().describe("Activity subject/title"),
    type: z.string().optional().describe("Activity type (call, meeting, task, deadline, email, lunch)"),
    due_date: z.string().optional().describe("Due date (YYYY-MM-DD)"),
    due_time: z.string().optional().describe("Due time (HH:MM)"),
    duration: z.string().optional().describe("Duration (HH:MM)"),
    deal_id: z.number().optional().describe("Deal ID"),
    person_id: z.number().optional().describe("Person ID"),
    org_id: z.number().optional().describe("Organization ID"),
    note: z.string().optional().describe("Activity note"),
    done: z.boolean().optional().describe("Whether activity is done")
  },
  async ({ activityId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await activitiesApi.updateActivity(activityId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating activity ${activityId}:`, error);
      return { content: [{ type: "text", text: `Error updating activity ${activityId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Add a lead
server.tool(
  "add-lead",
  "Create a new lead in Pipedrive",
  {
    title: z.string().describe("Lead title"),
    person_id: z.number().optional().describe("Person ID to link this lead to"),
    organization_id: z.number().optional().describe("Organization ID to link this lead to"),
    value: z.object({
      amount: z.number(),
      currency: z.string()
    }).optional().describe("Lead value e.g. { amount: 1000, currency: 'USD' }"),
    expected_close_date: z.string().optional().describe("Expected close date (YYYY-MM-DD)"),
    note: z.string().optional().describe("Note about the lead")
  },
  async (opts) => {
    try {
      // @ts-ignore
      const response = await leadsApi.addLead(opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error("Error adding lead:", error);
      return { content: [{ type: "text", text: `Error adding lead: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// Update a lead
server.tool(
  "update-lead",
  "Update an existing lead in Pipedrive",
  {
    leadId: z.string().describe("Pipedrive lead ID (UUID)"),
    title: z.string().optional().describe("Lead title"),
    person_id: z.number().optional().describe("Person ID"),
    organization_id: z.number().optional().describe("Organization ID"),
    value: z.object({
      amount: z.number(),
      currency: z.string()
    }).optional().describe("Lead value e.g. { amount: 1000, currency: 'USD' }"),
    expected_close_date: z.string().optional().describe("Expected close date (YYYY-MM-DD)"),
    is_archived: z.boolean().optional().describe("Whether to archive the lead")
  },
  async ({ leadId, ...opts }) => {
    try {
      // @ts-ignore
      const response = await leadsApi.updateLead(leadId, opts);
      return {
        content: [{ type: "text", text: JSON.stringify(response.data, null, 2) }]
      };
    } catch (error) {
      console.error(`Error updating lead ${leadId}:`, error);
      return { content: [{ type: "text", text: `Error updating lead ${leadId}: ${getErrorMessage(error)}` }], isError: true };
    }
  }
);

// === PROMPTS ===

// Prompt for getting all deals
server.prompt(
  "list-all-deals",
  "List all deals in Pipedrive",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please list all deals in my Pipedrive account, showing their title, value, status, and stage."
      }
    }]
  })
);

// Prompt for getting all persons
server.prompt(
  "list-all-persons",
  "List all persons in Pipedrive",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please list all persons in my Pipedrive account, showing their name, email, phone, and organization."
      }
    }]
  })
);

// Prompt for getting all pipelines
server.prompt(
  "list-all-pipelines",
  "List all pipelines in Pipedrive",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please list all pipelines in my Pipedrive account, showing their name and stages."
      }
    }]
  })
);

// Prompt for analyzing deals
server.prompt(
  "analyze-deals",
  "Analyze deals by stage",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please analyze the deals in my Pipedrive account, grouping them by stage and providing total value for each stage."
      }
    }]
  })
);

// Prompt for analyzing contacts
server.prompt(
  "analyze-contacts",
  "Analyze contacts by organization",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please analyze the persons in my Pipedrive account, grouping them by organization and providing a count for each organization."
      }
    }]
  })
);

// Prompt for analyzing leads
server.prompt(
  "analyze-leads",
  "Analyze leads by status",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please search for all leads in my Pipedrive account and group them by status."
      }
    }]
  })
);

// Prompt for pipeline comparison
server.prompt(
  "compare-pipelines",
  "Compare different pipelines and their stages",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please list all pipelines in my Pipedrive account and compare them by showing the stages in each pipeline."
      }
    }]
  })
);

// Prompt for finding high-value deals
server.prompt(
  "find-high-value-deals",
  "Find high-value deals",
  {},
  () => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: "Please identify the highest value deals in my Pipedrive account and provide information about which stage they're in and which person or organization they're associated with."
      }
    }]
  })
);

// Get transport type from environment variable (default to stdio)
const transportType = process.env.MCP_TRANSPORT || 'stdio';

if (transportType === 'sse') {
  // SSE transport - create HTTP server
  const port = parseInt(process.env.MCP_PORT || '3000', 10);
  const endpoint = process.env.MCP_ENDPOINT || '/message';

  // Store active transports by session ID
  const transports = new Map<string, SSEServerTransport>();

  const httpServer = http.createServer(async (req, res) => {
    const url = new URL(req.url!, `http://${req.headers.host}`);

    // Enable CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Session-Id');

    if (req.method === 'OPTIONS') {
      res.writeHead(204);
      res.end();
      return;
    }

    if (req.method === 'GET' && url.pathname === '/sse') {
      const authResult = verifyRequestAuthentication(req);
      if (!authResult.ok) {
        res.writeHead(authResult.status, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: authResult.message }));
        return;
      }

      // Establish SSE connection
      console.error('New SSE connection request');
      const transport = new SSEServerTransport(endpoint, res);

      // Store transport by session ID
      transports.set(transport.sessionId, transport);

      transport.onclose = () => {
        console.error(`SSE connection closed: ${transport.sessionId}`);
        transports.delete(transport.sessionId);
      };

      try {
        await server.connect(transport);
        console.error(`SSE connection established: ${transport.sessionId}`);
      } catch (err) {
        console.error('Failed to establish SSE connection:', err);
        transports.delete(transport.sessionId);
      }
    } else if (req.method === 'POST' && url.pathname === endpoint) {
      const authResult = verifyRequestAuthentication(req);
      if (!authResult.ok) {
        res.writeHead(authResult.status, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: authResult.message }));
        return;
      }

      // Handle incoming message
      const sessionId = url.searchParams.get('sessionId') || req.headers['x-session-id'] as string;

      if (!sessionId) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Missing sessionId' }));
        return;
      }

      const transport = transports.get(sessionId);
      if (!transport) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Session not found' }));
        return;
      }

      req.on('error', err => {
        console.error('Error receiving POST message body:', err);
        if (!res.headersSent) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Invalid request body' }));
        }
      });

      try {
        await transport.handlePostMessage(req, res);
      } catch (err) {
        console.error('Error handling POST message:', err);
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Internal server error' }));
        }
      }
    } else {
      // Health check endpoint
      if (req.method === 'GET' && url.pathname === '/health') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ status: 'ok', transport: 'sse' }));
        return;
      }

      res.writeHead(404);
      res.end('Not found');
    }
  });

  httpServer.listen(port, () => {
    console.error(`Pipedrive MCP Server (SSE) listening on port ${port}`);
    console.error(`SSE endpoint: http://localhost:${port}/sse`);
    console.error(`Message endpoint: http://localhost:${port}${endpoint}`);
  });
} else {
  // Default: stdio transport
  const transport = new StdioServerTransport();
  server.connect(transport).catch(err => {
    console.error("Failed to start MCP server:", err);
    process.exit(1);
  });

  console.error("Pipedrive MCP Server started (stdio transport)");
}
