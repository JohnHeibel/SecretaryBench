import { NextRequest, NextResponse } from "next/server";

// Single shared passcode gate. If APP_PASSCODE is unset, auth is off in local dev —
// but on a deployment that would be wide open, so we FAIL CLOSED there instead of
// silently disabling the gate. ~20 trusted users — deliberately lightweight.
export function middleware(req: NextRequest) {
  const pass = process.env.APP_PASSCODE;
  if (!pass) {
    if (process.env.VERCEL_ENV) {
      return new NextResponse("APP_PASSCODE is not configured for this deployment.", { status: 503 });
    }
    return NextResponse.next(); // local dev only
  }

  const { pathname } = req.nextUrl;
  if (pathname.startsWith("/login") || pathname.startsWith("/api/auth")) return NextResponse.next();

  if (req.cookies.get("sb_pass")?.value === pass) return NextResponse.next();

  const url = req.nextUrl.clone();
  url.pathname = "/login";
  return NextResponse.redirect(url);
}

export const config = {
  // gate everything except Next internals and static assets
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
