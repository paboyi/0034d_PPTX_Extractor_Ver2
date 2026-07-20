import serverlessHttp from "serverless-http";
import { app } from "./server";

export const handler = serverlessHttp(app);

// This is the actual entry point AWS will call.
// wraps our Express app so Lambda can invoke it like a function instead of a running server.