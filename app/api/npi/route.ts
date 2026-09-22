import { NextRequest } from "next/server";

export const dynamic = "force-dynamic";

function clean(value: string | null, maxLength: number) {
  return String(value ?? "").trim().slice(0, maxLength);
}

export async function GET(request: NextRequest) {
  const number = clean(request.nextUrl.searchParams.get("number"), 10);
  const firstName = clean(request.nextUrl.searchParams.get("first_name"), 70);
  const lastName = clean(request.nextUrl.searchParams.get("last_name"), 70);
  const state = clean(request.nextUrl.searchParams.get("state"), 2).toUpperCase();
  if ((number && !/^\d{10}$/.test(number)) || (!number && !firstName && !lastName) || (state && !/^[A-Z]{2}$/.test(state))) {
    return Response.json({ error: "Supply a 10-digit number or at least one name field; state must be two letters when supplied" }, { status: 400 });
  }

  const params = new URLSearchParams({ version: "2.1" });
  if (number) {
    params.set("number", number);
  } else {
    params.set("enumeration_type", "NPI-1");
    if (firstName) params.set("first_name", firstName);
    if (lastName) params.set("last_name", lastName);
    params.set("limit", "20");
    if (state) params.set("state", state);
  }
  try {
    const upstream = await fetch(`https://npiregistry.cms.hhs.gov/api/?${params.toString()}`, {
      headers: { Accept: "application/json" }, cache: "no-store",
    });
    if (!upstream.ok) return Response.json({ error: "NPPES request failed" }, { status: upstream.status });
    const data = await upstream.json();
    return Response.json(data, { headers: { "Cache-Control": "private, no-store", "X-Content-Type-Options": "nosniff" } });
  } catch {
    return Response.json({ error: "NPPES is temporarily unavailable" }, { status: 502 });
  }
}
