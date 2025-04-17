import { cookies } from 'next/headers';
import { NextResponse } from 'next/server';
import jwt from 'jsonwebtoken';

const JWT_SECRET = process.env.JWT_SECRET!;

interface UserPayload {
  id: string;
  email: string;
  role: string;
}

type AuthenticatedRouteHandler = (
  request: Request,
) => Promise<NextResponse> | NextResponse;


export function withAuthRoute(handler: AuthenticatedRouteHandler) {
  return async (request: Request): Promise<NextResponse> => {
    const cookieStore = cookies();
    const token = cookieStore.get('accessToken')?.value;

    if (!token) {
       return NextResponse.json({ message: 'Authentication required: No token found.' }, { status: 401 });
    }

    try {
      const decoded = jwt.verify(token, JWT_SECRET) as UserPayload;
      return await handler(request);

    } catch (error) {
      if (error instanceof jwt.TokenExpiredError) {
        return NextResponse.json({ message: 'Authentication required: Token expired.', code: 'TOKEN_EXPIRED' }, { status: 401 });
      } else if (error instanceof jwt.JsonWebTokenError) {
         return NextResponse.json({ message: 'Authentication required: Invalid token.' }, { status: 401 });
      } else {
        console.error("Auth Middleware Error:", error);
        return NextResponse.json({ message: 'Internal server error during authentication.' }, { status: 500 });
      }
    }
  };
}
