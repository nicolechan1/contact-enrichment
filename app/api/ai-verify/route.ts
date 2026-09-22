import { NextRequest } from "next/server";

export const dynamic = "force-dynamic";

type VerificationInput = {
  name: string;
  npi: string;
  npiSpecialty: string;
  claimedSpecialty: string;
  organization: string;
  organizationDomain: string;
  state: string;
  country: string;
  linkedInUrl: string;
  practiceLocation: string;
  verifyUsLocation: boolean;
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

export async function GET() {
  return Response.json(
    { configured: Boolean(process.env.OPENAI_API_KEY) },
    { headers: { "Cache-Control": "private, no-store", "X-Content-Type-Options": "nosniff" } },
  );
}

export async function POST(request: NextRequest) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return Response.json({ error: "AI verification is not configured." }, { status: 503 });
  }

  let body: Partial<VerificationInput>;
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON request." }, { status: 400 });
  }

  const input: VerificationInput = {
    name: clean(body.name, 160),
    npi: clean(body.npi, 10),
    npiSpecialty: clean(body.npiSpecialty, 180),
    claimedSpecialty: clean(body.claimedSpecialty, 180),
    organization: clean(body.organization, 180),
    organizationDomain: clean(body.organizationDomain, 120),
    state: clean(body.state, 40),
    country: clean(body.country, 80),
    linkedInUrl: clean(body.linkedInUrl, 500),
    practiceLocation: clean(body.practiceLocation, 240),
    verifyUsLocation: body.verifyUsLocation !== false,
  };

  if (!input.name || !/^\d{10}$/.test(input.npi)) {
    return Response.json({ error: "A name and 10-digit NPI are required." }, { status: 400 });
  }

  const prompt = `Research this healthcare professional using current public web sources. NPPES is the source of the NPI; do not propose or change an NPI.

Contact data:
${JSON.stringify(input, null, 2)}

Make separate decisions about identity, physician status, current U.S. practice, and specialty match. ${input.verifyUsLocation ? "Independently verify current U.S. practice from public web evidence." : "The source already supplies United States; do not spend web-search effort rechecking geography. Set practices_in_us true from the supplied source and focus research on identity, physician status, and specialty."} Do not treat a different person, an identity conflict, a Ph.D., or missing evidence as proof that someone is outside the United States. Use null when a researched yes/no decision cannot be supported. Return us_doctor_confidence from 0 to 100 for the combined conclusion. Return specialty_match_score from 0 to 100 when claimed and verified specialties can be compared, otherwise null. Prefer hospital, university, clinic, state licensing, board-certification, and professional-association sources. A public LinkedIn page may support identity but must not be the only evidence. Use one targeted search combining name, organization, NPI, and specialty; only search again if identity remains unclear. Each reason must be one plain sentence under 180 characters; put details in evidence URLs.`;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 35_000);
  try {
    const upstream = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      signal: controller.signal,
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-5.6-luna",
        tools: [{ type: "web_search", search_context_size: "low" }],
        tool_choice: "required",
        input: prompt,
        text: {
          format: {
            type: "json_schema",
            name: "doctor_verification",
            strict: true,
            schema: {
              type: "object",
              properties: {
                identity_match: { type: ["boolean", "null"] },
                practices_in_us: { type: ["boolean", "null"] },
                us_practice_reason: { type: "string" },
                is_physician: { type: ["boolean", "null"] },
                us_doctor_confidence: { type: "integer", minimum: 0, maximum: 100 },
                verified_specialty: { type: "string" },
                specialty_match_score: { type: ["integer", "null"], minimum: 0, maximum: 100 },
                specialty_match_reason: { type: "string" },
                confidence: { type: "string", enum: ["high", "medium", "low", "insufficient"] },
                reason: { type: "string" },
                evidence_urls: { type: "array", items: { type: "string" }, maxItems: 5 },
              },
              required: ["identity_match", "practices_in_us", "us_practice_reason", "is_physician", "us_doctor_confidence", "verified_specialty", "specialty_match_score", "specialty_match_reason", "confidence", "reason", "evidence_urls"],
              additionalProperties: false,
            },
          },
        },
      }),
    });

    if (!upstream.ok) {
      return Response.json(
        { error: "OpenAI verification request failed.", detail: await openAiErrorMessage(upstream) },
        { status: upstream.status >= 400 && upstream.status < 500 ? upstream.status : 502 },
      );
    }

    const payload = await upstream.json();
    const text = responseText(payload);
    const parsed = JSON.parse(text) as {
      identity_match: boolean | null;
      practices_in_us: boolean | null;
      us_practice_reason: string;
      is_physician: boolean | null;
      us_doctor_confidence: number;
      verified_specialty: string;
      specialty_match_score: number | null;
      specialty_match_reason: string;
      confidence: "high" | "medium" | "low" | "insufficient";
      reason: string;
      evidence_urls: string[];
    };

    return Response.json({
      identityMatch: typeof parsed.identity_match === "boolean" ? parsed.identity_match : null,
      practicesInUs: typeof parsed.practices_in_us === "boolean" ? parsed.practices_in_us : null,
      usPracticeReason: clean(parsed.us_practice_reason, 220),
      isPhysician: typeof parsed.is_physician === "boolean" ? parsed.is_physician : null,
      usDoctorConfidence: Number.isInteger(parsed.us_doctor_confidence) ? Math.max(0, Math.min(100, parsed.us_doctor_confidence)) : 0,
      verifiedSpecialty: clean(parsed.verified_specialty, 180),
      specialtyMatchScore: Number.isInteger(parsed.specialty_match_score) ? Math.max(0, Math.min(100, parsed.specialty_match_score as number)) : null,
      specialtyMatchReason: clean(parsed.specialty_match_reason, 220),
      confidence: ["high", "medium", "low", "insufficient"].includes(parsed.confidence) ? parsed.confidence : "insufficient",
      reason: clean(parsed.reason, 220),
      evidenceUrls: safeEvidenceUrls(parsed.evidence_urls),
    }, { headers: { "Cache-Control": "private, no-store", "X-Content-Type-Options": "nosniff" } });
  } catch (error) {
    return Response.json({
      error: controller.signal.aborted ? "AI verification timed out; continue with the next contact." : "AI verification is temporarily unavailable.",
      detail: clean(error instanceof Error ? error.message : "Unknown response-processing error.", 500),
    }, { status: 502 });
  } finally {
    clearTimeout(timeout);
  }
}
