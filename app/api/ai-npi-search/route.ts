import { NextRequest } from "next/server";

export const dynamic = "force-dynamic";

type SearchInput = {
  name: string;
  organization: string;
  organizationDomain: string;
  state: string;
  country: string;
  claimedSpecialty: string;
  linkedInUrl: string;
  npiCandidates: Array<{
    number: string;
    registryName: string;
    specialty: string;
    city: string;
    state: string;
  }>;
};

function clean(value: unknown, maxLength: number) {
  return String(value ?? "").trim().slice(0, maxLength);
}

function responseText(payload: unknown) {
  if (!payload || typeof payload !== "object") return "";
  const response = payload as { output_text?: string; output?: Array<{ content?: Array<{ type?: string; text?: string }> }> };
  if (response.output_text) return response.output_text;
  return (response.output ?? [])
    .flatMap((item) => item.content ?? [])
    .filter((item) => item.type === "output_text")
    .map((item) => item.text ?? "")
    .join("");
}

function safeEvidenceUrls(value: unknown) {
  if (!Array.isArray(value)) return [];
  return value.slice(0, 5).flatMap((item) => {
    try {
      const url = new URL(String(item));
      return ["http:", "https:"].includes(url.protocol) ? [url.toString()] : [];
    } catch {
      return [];
    }
  });
}

async function openAiErrorMessage(response: Response) {
  try {
    const payload = await response.json() as { error?: { message?: string } | string };
    const message = typeof payload.error === "string" ? payload.error : payload.error?.message;
    return clean(message, 500) || `OpenAI returned HTTP ${response.status}.`;
  } catch {
    return `OpenAI returned HTTP ${response.status}.`;
  }
}

export async function POST(request: NextRequest) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) return Response.json({ error: "AI NPI search is not configured." }, { status: 503 });

  let body: Partial<SearchInput>;
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON request." }, { status: 400 });
  }

  const input: SearchInput = {
    name: clean(body.name, 160),
    organization: clean(body.organization, 180),
    organizationDomain: clean(body.organizationDomain, 120),
    state: clean(body.state, 40),
    country: clean(body.country, 80),
    claimedSpecialty: clean(body.claimedSpecialty, 180),
    linkedInUrl: clean(body.linkedInUrl, 500),
    npiCandidates: Array.isArray(body.npiCandidates) ? body.npiCandidates.slice(0, 20).map((candidate) => ({
      number: clean(candidate?.number, 10),
      registryName: clean(candidate?.registryName, 180),
      specialty: clean(candidate?.specialty, 180),
      city: clean(candidate?.city, 100),
      state: clean(candidate?.state, 20),
    })) : [],
  };
  if (!input.name) return Response.json({ error: "A contact name is required." }, { status: 400 });

  const prompt = `Research this healthcare professional using current public web sources because an exact name search in NPPES returned no results.

Contact data:
${JSON.stringify(input, null, 2)}

The contact's work-email domain may identify their institution, but the email address itself has not been shared. The NPI candidates are broad first-name matches returned directly by NPPES; use them as a shortlist, not as proof of identity.

Look for a possible U.S. National Provider Identifier and for legal, married, maiden, compound, or professional-name variations. Prefer official NPPES, hospital, university, clinic, state licensing, board-certification, and professional-association sources. A public LinkedIn page may support identity but must not be the only evidence. Do not guess an NPI. Return a candidate only when public evidence connects this exact person to the number with high or medium confidence. Provide the exact first and last name shown for that number. Prefer a candidate from the supplied NPPES shortlist when evidence supports it. If the identity link is weak or no number is found, return an empty candidate_npi and insufficient confidence. The application will independently verify every candidate against NPPES before accepting it. Use one targeted search combining name, organization, state, and specialty; only search again if an alternate legal name is strongly indicated. Keep reason to one plain sentence under 180 characters; put details in evidence URLs.`;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 35_000);
  try {
    const upstream = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-5.6-luna",
        tools: [{ type: "web_search", search_context_size: "low" }],
        tool_choice: "required",
        input: prompt,
        text: {
          format: {
            type: "json_schema",
            name: "npi_discovery",
            strict: true,
            schema: {
              type: "object",
              properties: {
                candidate_npi: { type: "string" },
                registry_first_name: { type: "string" },
                registry_last_name: { type: "string" },
                alternate_name: { type: "string" },
                confidence: { type: "string", enum: ["high", "medium", "low", "insufficient"] },
                reason: { type: "string" },
                evidence_urls: { type: "array", items: { type: "string" }, maxItems: 5 },
              },
              required: ["candidate_npi", "registry_first_name", "registry_last_name", "alternate_name", "confidence", "reason", "evidence_urls"],
              additionalProperties: false,
            },
          },
        },
      }),
    });
    if (!upstream.ok) {
      return Response.json(
        { error: "OpenAI NPI search failed.", detail: await openAiErrorMessage(upstream) },
        { status: upstream.status >= 400 && upstream.status < 500 ? upstream.status : 502 },
      );
    }

    const payload = await upstream.json();
    const parsed = JSON.parse(responseText(payload)) as {
      candidate_npi: string;
      registry_first_name: string;
      registry_last_name: string;
      alternate_name: string;
      confidence: "high" | "medium" | "low" | "insufficient";
      reason: string;
      evidence_urls: string[];
    };
    const candidateNpi = clean(parsed.candidate_npi, 30).replace(/\D/g, "").slice(0, 10);
    return Response.json({
      candidateNpi: /^\d{10}$/.test(candidateNpi) ? candidateNpi : "",
      registryFirstName: clean(parsed.registry_first_name, 100),
      registryLastName: clean(parsed.registry_last_name, 100),
      alternateName: clean(parsed.alternate_name, 180),
      confidence: ["high", "medium", "low", "insufficient"].includes(parsed.confidence) ? parsed.confidence : "insufficient",
      reason: clean(parsed.reason, 220),
      evidenceUrls: safeEvidenceUrls(parsed.evidence_urls),
    }, { headers: { "Cache-Control": "private, no-store", "X-Content-Type-Options": "nosniff" } });
  } catch (error) {
    return Response.json({
      error: controller.signal.aborted ? "AI NPI search timed out; continue with the next contact." : "AI NPI search is temporarily unavailable.",
      detail: clean(error instanceof Error ? error.message : "Unknown response-processing error.", 500),
    }, { status: 502 });
  } finally {
    clearTimeout(timeout);
  }
}
