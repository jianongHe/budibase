const mockSearchGeminiFileStore = jest.fn()
const mockIngestGeminiFile = jest.fn()
const mockDeleteGeminiFileFromStore = jest.fn()
const mockUpdateKnowledgeBaseFile = jest.fn()

jest.mock("../../knowledgeBase/geminiFileStore", () => ({
  searchGeminiFileStore: (...args: any[]) => mockSearchGeminiFileStore(...args),
  ingestGeminiFile: (...args: any[]) => mockIngestGeminiFile(...args),
  deleteGeminiFileFromStore: (...args: any[]) =>
    mockDeleteGeminiFileFromStore(...args),
}))

jest.mock("../../knowledgeBase", () => ({
  updateKnowledgeBaseFile: (...args: any[]) =>
    mockUpdateKnowledgeBaseFile(...args),
}))

import { KnowledgeBase, KnowledgeBaseType } from "@budibase/types"
import * as XLSX from "xlsx"
import { GeminiRagProcessor } from "./gemini"

const knowledgeBase = {
  _id: "kb_1",
  name: "KB",
  type: KnowledgeBaseType.GEMINI,
  config: {
    googleFileStoreId: "store_1",
  },
} satisfies KnowledgeBase

const buildWorkbookBuffer = (
  rows: Array<Array<string | number>>,
  sheetName = "Plans"
) => {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows),
    sheetName
  )
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" })
}

describe("GeminiRagProcessor", () => {
  beforeEach(() => {
    jest.clearAllMocks()
  })

  it("maps search sourceId from Gemini file_id", async () => {
    mockSearchGeminiFileStore.mockResolvedValue([
      {
        file_id: "gemini-file-1",
        filename: "policy.md",
        score: 0.9,
        content: [{ type: "text", text: "4-day policy" }],
      },
    ])

    const processor = new GeminiRagProcessor(knowledgeBase)

    const result = await processor.search("What is policy?")

    expect(result).toEqual([
      {
        source: "gemini-file-1",
        chunkText: "4-day policy",
      },
    ])
  })

  it("falls back to filename when Gemini file_id is missing", async () => {
    mockSearchGeminiFileStore.mockResolvedValue([
      {
        file_id: null,
        filename: "policy.md",
        score: 0.9,
        content: [{ type: "text", text: "4-day policy" }],
      },
    ])

    const processor = new GeminiRagProcessor(knowledgeBase)

    const result = await processor.search("What is policy?")

    expect(result).toEqual([
      {
        source: "policy.md",
        chunkText: "4-day policy",
      },
    ])
  })

  it("normalizes spreadsheet uploads into plain text before ingestion", async () => {
    mockIngestGeminiFile.mockResolvedValue({ fileId: "gemini-file-1" })

    const processor = new GeminiRagProcessor(knowledgeBase)

    const file = {
      _id: "file_1",
      filename: "pricing.xlsx",
      mimetype:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      status: "processing",
    } as any

    await processor.ingestKnowledgeBaseFile(
      file,
      buildWorkbookBuffer([
        ["Plan", "Price"],
        ["Pro", 49],
      ])
    )

    expect(mockIngestGeminiFile).toHaveBeenCalledTimes(1)
    const ingestPayload = mockIngestGeminiFile.mock.calls[0][0]
    expect(ingestPayload.vectorStoreId).toBe("store_1")
    expect(ingestPayload.filename).toBe("pricing.xlsx")
    expect(ingestPayload.mimetype).toBe("text/plain")
    expect(ingestPayload.buffer.toString("utf8")).toContain(
      "Columns: Plan | Price"
    )
    expect(ingestPayload.buffer.toString("utf8")).toContain("Plan: Pro")
    expect(mockUpdateKnowledgeBaseFile).toHaveBeenCalledWith(
      expect.objectContaining({
        status: "ready",
        ragSourceId: "gemini-file-1",
      })
    )
  })
})
