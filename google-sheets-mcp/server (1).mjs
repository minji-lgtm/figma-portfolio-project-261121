import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import { google } from "googleapis";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({
  path: path.join(__dirname, ".env"),
});

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const DEFAULT_SPREADSHEET_ID = process.env.GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID;

if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  throw new Error("GOOGLE_APPLICATION_CREDENTIALS is not set.");
}

const credentialsPath = path.resolve(
  __dirname,
  process.env.GOOGLE_APPLICATION_CREDENTIALS,
);

const auth = new google.auth.GoogleAuth({
  keyFile: credentialsPath,
  scopes: SCOPES,
});

const sheets = google.sheets({ version: "v4", auth });

function ensureSpreadsheetId(spreadsheetId) {
  const id = spreadsheetId || DEFAULT_SPREADSHEET_ID;
  if (!id) {
    throw new Error(
      "spreadsheetId is required or set GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID.",
    );
  }
  return id;
}

function textResult(obj) {
  return {
    content: [
      {
        type: "text",
        text: typeof obj === "string" ? obj : JSON.stringify(obj, null, 2),
      },
    ],
  };
}

const server = new McpServer({
  name: "google-sheets-mcp",
  version: "1.0.0",
});

// mcp 서버가 GS에 접속해서 데이터 읽어오기
server.registerTool(
  "gsheets_read_range",
  {
    title: "Read a range from Google Sheets",
    description: "Reads values from a Google Sheet range in A1 notation.",
    inputSchema: {
      spreadsheetId: z.string().optional(),
      range: z.string(),
    },
  },
  async ({ spreadsheetId, range }) => {
    try {
      const id = ensureSpreadsheetId(spreadsheetId);
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
      });

      return textResult({
        spreadsheetId: id,
        range,
        values: res.data.values || [],
      });
    } catch (error) {
      return textResult({
        error: error instanceof Error ? error.message : String(error),
      });
    }
  },
);

server.registerTool(
  "gsheets_append_rows",
  {
    title: "Append rows to Google Sheets",
    description: "Appends one or more rows to the end of a sheet range.",
    inputSchema: {
      spreadsheetId: z.string().optional(),
      range: z.string(),
      values: z.array(z.array(z.union([z.string(), z.number(), z.boolean()]))),
    },
  },
  async ({ spreadsheetId, range, values }) => {
    try {
      const id = ensureSpreadsheetId(spreadsheetId);
      const res = await sheets.spreadsheets.values.append({
        spreadsheetId: id,
        range,
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values },
      });

      return textResult({
        spreadsheetId: id,
        range,
        updates: res.data.updates,
      });
    } catch (error) {
      return textResult({
        error: error instanceof Error ? error.message : String(error),
      });
    }
  },
);

server.registerTool(
  "gsheets_update_range",
  {
    title: "Update a range in Google Sheets",
    description: "Overwrites a specific range with the provided values.",
    inputSchema: {
      spreadsheetId: z.string().optional(),
      range: z.string(),
      values: z.array(z.array(z.union([z.string(), z.number(), z.boolean()]))),
    },
  },
  async ({ spreadsheetId, range, values }) => {
    try {
      const id = ensureSpreadsheetId(spreadsheetId);
      const res = await sheets.spreadsheets.values.update({
        spreadsheetId: id,
        range,
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });

      return textResult({
        spreadsheetId: id,
        range,
        updatedRows: res.data.updatedRows,
        updatedColumns: res.data.updatedColumns,
        updatedCells: res.data.updatedCells,
      });
    } catch (error) {
      return textResult({
        error: error instanceof Error ? error.message : String(error),
      });
    }
  },
);

server.registerTool(
  "gsheets_clear_range",
  {
    title: "Clear a range in Google Sheets",
    description: "Clears values from a Google Sheet range.",
    inputSchema: {
      spreadsheetId: z.string().optional(),
      range: z.string(),
    },
  },
  async ({ spreadsheetId, range }) => {
    try {
      const id = ensureSpreadsheetId(spreadsheetId);
      const res = await sheets.spreadsheets.values.clear({
        spreadsheetId: id,
        range,
        requestBody: {},
      });

      return textResult({
        spreadsheetId: id,
        clearedRange: res.data.clearedRange,
      });
    } catch (error) {
      return textResult({
        error: error instanceof Error ? error.message : String(error),
      });
    }
  },
);

server.registerTool(
  "gsheets_write_keyword_summary",
  {
    title: "Write keyword summary rows",
    description:
      "Writes keyword summary rows with columns like keyword, count, source, date.",
    inputSchema: {
      spreadsheetId: z.string().optional(),
      sheetName: z.string().default("KeywordSummary"),
      rows: z.array(
        z.object({
          keyword: z.string(),
          count: z.number(),
          source: z.string().default("slack"),
          date: z.string(),
        }),
      ),
    },
  },
  async ({ spreadsheetId, sheetName, rows }) => {
    try {
      const id = ensureSpreadsheetId(spreadsheetId);
      const values = rows.map((row) => [
        row.keyword,
        row.count,
        row.source,
        row.date,
      ]);

      const range = `${sheetName}!A:D`;

      const res = await sheets.spreadsheets.values.append({
        spreadsheetId: id,
        range,
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values },
      });

      return textResult({
        spreadsheetId: id,
        sheetName,
        appended: rows.length,
        updates: res.data.updates,
      });
    } catch (error) {
      return textResult({
        error: error instanceof Error ? error.message : String(error),
      });
    }
  },
);

const transport = new StdioServerTransport();

await server.connect(transport);
