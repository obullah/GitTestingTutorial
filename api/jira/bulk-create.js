// api/jira/bulk-create.js
import axios from "axios";

/**
 * Build ADF (Atlassian Document Format) for a plain-text description.
 * This keeps your Excel "Description" column simple, but satisfies Jira Cloud.
 */
function buildAdfDescription(text) {
  const safeText = (text || "").toString().trim();

  return {
    type: "doc",
    version: 1,
    content: safeText
      ? [
          {
            type: "paragraph",
            content: [
              {
                type: "text",
                text: safeText,
              },
            ],
          },
        ]
      : [],
  };
}

/**
 * Resolve "assignee" text (email, display name, or accountId) to a Jira accountId.
 * - If it's already an accountId, we just use it.
 * - Otherwise, we search Jira users with ?query=<value> and grab the first match.
 * hkhjkjhoihiohhngutfrdesrdtfyguihiiugfdszxfcghvjbn
 */
async function resolveAssigneeAccountId(jiraClient, assigneeRaw) {
  if (!assigneeRaw) return null;

  const value = assigneeRaw.toString().trim();
  if (!value) return null;

  // Heuristic: if it "looks like" an Atlassian accountId, just use it as-is.
  // (Typical accountIds are long, opaque strings.)
  if (value.length >= 20 && !value.includes("@")) {
    return value;
  }

  // Otherwise, search by query (works for display name or email if allowed).
  const searchRes = await jiraClient.get("/user/search", {
    params: { query: value, maxResults: 1 },
  });

  const users = searchRes.data || [];
  if (!users.length) {
    throw new Error(
      `Could not find Jira user matching "${value}". Check the Assignee value.`
    );
  }

  return users[0].accountId;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);
    return res.status(405).json({ message: "Method Not Allowed" });
  }

  const {
    JIRA_BASE_URL,
    JIRA_EMAIL,
    JIRA_API_TOKEN,
    JIRA_SPACE_KEY,
    JIRA_ISSUE_TYPE = "Story",
    JIRA_STORY_POINTS_FIELD = "customfield_10016",
  } = process.env;

  if (!JIRA_BASE_URL || !JIRA_EMAIL || !JIRA_API_TOKEN || !JIRA_SPACE_KEY) {
    return res.status(500).json({
      message:
        "Missing Jira env vars. Set JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN, JIRA_SPACE_KEY in Vercel.",
    });
  }

  const { issues } = req.body || {};
  if (!Array.isArray(issues) || issues.length === 0) {
    return res
      .status(400)
      .json({ message: "Body must contain a non-empty 'issues' array." });
  }

  const jiraClient = axios.create({
    baseURL: `${JIRA_BASE_URL}/rest/api/3`,
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      Authorization:
        "Basic " +
        Buffer.from(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`).toString("base64"),
    },
  });

  const results = [];

  for (const issue of issues) {
    const {
      excelRowIndex,
      issueId,
      assignee,
      storyPoints,
      summary,
      description,
    } = issue;

    if (!summary) {
      results.push({
        excelRowIndex,
        issueId,
        success: false,
        errorMessage: "Missing summary – issue not created.",
      });
      continue;
    }

    // Build Jira fields
    const fields = {
      project: { key: JIRA_SPACE_KEY },
      summary: summary.toString(),
      description: buildAdfDescription(description),
      issuetype: { name: JIRA_ISSUE_TYPE },
    };

    if (
      storyPoints !== null &&
      storyPoints !== undefined &&
      storyPoints !== ""
    ) {
      fields[JIRA_STORY_POINTS_FIELD] = Number(storyPoints);
    }

    // Resolve assignee → accountId
    if (assignee) {
      try {
        const accountId = await resolveAssigneeAccountId(jiraClient, assignee);
        if (accountId) {
          fields.assignee = { accountId };
        } else {
          // If no accountId found, mark this row as failed and skip create.
          results.push({
            excelRowIndex,
            issueId,
            jiraKey: null,
            success: false,
            errorMessage: `Assignee "${assignee}" could not be resolved to a Jira user.`,
          });
          continue;
        }
      } catch (err) {
        results.push({
          excelRowIndex,
          issueId,
          jiraKey: null,
          success: false,
          errorMessage:
            err.message ||
            `Failed to resolve assignee "${assignee}". Check the value in Excel.`,
        });
        continue;
      }
    }

    try {
      const response = await jiraClient.post("/issue", { fields });
      results.push({
        excelRowIndex,
        issueId,
        jiraKey: response.data.key,
        success: true,
        errorMessage: null,
      });
    } catch (err) {
      console.error("Jira create error:", err?.response?.data || err.message);
      const message =
        err?.response?.data?.errors
          ? JSON.stringify(err.response.data.errors)
          : err?.response?.data?.errorMessages?.join(", ") ||
            err.message ||
            "Unknown error";

      results.push({
        excelRowIndex,
        issueId,
        jiraKey: null,
        success: false,
        errorMessage: message,
      });
    }
  }

  const successCount = results.filter((r) => r.success).length;
  const failureCount = results.length - successCount;

  return res.status(200).json({
    successCount,
    failureCount,
    results,
  });
}
