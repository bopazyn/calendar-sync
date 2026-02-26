import { createHash, randomBytes } from "node:crypto";

export type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope?: string;
};

export const toBase64Url = (buffer: Buffer) => buffer
  .toString("base64")
  .replace(/\+/g, "-")
  .replace(/\//g, "_")
  .replace(/=+$/g, "");

export const createPkce = () => {
  const verifier = toBase64Url(randomBytes(32));
  const challenge = toBase64Url(createHash("sha256").update(verifier).digest());
  return { verifier, challenge };
};

export const delay = (ms: number) =>
  new Promise((resolve) => setTimeout(resolve, ms));