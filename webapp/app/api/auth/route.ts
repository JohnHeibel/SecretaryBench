import { NextResponse } from "next/server";

export async function POST(req: Request) {
  const { passcode } = (await req.json()) as { passcode?: string };
  if (!process.env.APP_PASSCODE || passcode !== process.env.APP_PASSCODE) {
    return NextResponse.json({ ok: false }, { status: 401 });
  }
  const res = NextResponse.json({ ok: true });
  res.cookies.set("sb_pass", passcode, { httpOnly: true, sameSite: "lax", secure: process.env.NODE_ENV === "production", path: "/", maxAge: 60 * 60 * 24 * 30 });
  return res;
}
