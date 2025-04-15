import { NextResponse } from "next/server";
import { getToken } from "next-auth/jwt";
import { NextRequest } from "next/server";

interface UserToken {
  user?: {
    role?: string;
  };
}

export async function middleware(req: NextRequest): Promise<NextResponse> {
  const token = (await getToken({
    req,
    secret: process.env.NEXTAUTH_SECRET,
  })) as UserToken | null;
  const url = req.nextUrl.clone();

  if (
    !token &&
    (!(url.pathname === "/") ||
      !url.pathname.startsWith("/login") ||
      !url.pathname.startsWith("/create"))
  ) {
    url.pathname = "/login";
    return NextResponse.redirect(url);
  }

  // Example: Restrict access to `/admin` for non-admin users
  if (!url.pathname.startsWith(`/${token?.user?.role}`)) {
    url.pathname = `/${token?.user?.role}`;
    return NextResponse.redirect(url);
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/dashboard/:path*", "/admin/:path*"],
};
