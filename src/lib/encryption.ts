const MAGIC = new TextEncoder().encode("PSTUDIOSEC1\n");
const ITERATIONS = 250_000;

interface SecureHeader {
  schema: "presentation-studio/encrypted-package";
  version: 1;
  algorithm: "AES-256-GCM";
  keyDerivation: "PBKDF2-SHA-256";
  iterations: number;
  salt: string;
  iv: string;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function concatBytes(...parts: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((sum, part) => sum + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

async function deriveKey(password: string, salt: Uint8Array, iterations: number): Promise<CryptoKey> {
  const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", hash: "SHA-256", salt: toArrayBuffer(salt), iterations },
    material,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"],
  );
}

function requirePassword(password: string): void {
  if (password.length < 12) throw new Error("Encrypted projects require a password of at least 12 characters.");
}

export function isEncryptedProject(bytes: Uint8Array): boolean {
  return MAGIC.every((byte, index) => bytes[index] === byte);
}

export async function encryptProjectPackage(plainBytes: Uint8Array, password: string): Promise<Uint8Array> {
  requirePassword(password);
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const header: SecureHeader = {
    schema: "presentation-studio/encrypted-package",
    version: 1,
    algorithm: "AES-256-GCM",
    keyDerivation: "PBKDF2-SHA-256",
    iterations: ITERATIONS,
    salt: bytesToBase64(salt),
    iv: bytesToBase64(iv),
  };
  const headerBytes = new TextEncoder().encode(JSON.stringify(header));
  const lengthBytes = new Uint8Array(4);
  new DataView(lengthBytes.buffer).setUint32(0, headerBytes.byteLength, false);
  const aad = concatBytes(MAGIC, lengthBytes, headerBytes);
  const key = await deriveKey(password, salt, ITERATIONS);
  const stablePlaintext = new Uint8Array(plainBytes.byteLength);
  stablePlaintext.set(plainBytes);
  const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv: toArrayBuffer(iv), additionalData: toArrayBuffer(aad), tagLength: 128 }, key, stablePlaintext.buffer));
  return concatBytes(aad, ciphertext);
}

export async function decryptProjectPackage(secureBytes: Uint8Array, password: string): Promise<Uint8Array> {
  requirePassword(password);
  if (!isEncryptedProject(secureBytes) || secureBytes.byteLength < MAGIC.byteLength + 4) throw new Error("This is not a Presentation Studio encrypted project.");
  const lengthOffset = MAGIC.byteLength;
  const headerLength = new DataView(secureBytes.buffer, secureBytes.byteOffset + lengthOffset, 4).getUint32(0, false);
  if (headerLength < 40 || headerLength > 16_384) throw new Error("The encrypted project header is invalid.");
  const headerStart = lengthOffset + 4;
  const cipherStart = headerStart + headerLength;
  if (cipherStart >= secureBytes.byteLength) throw new Error("The encrypted project is truncated.");
  const header = JSON.parse(new TextDecoder().decode(secureBytes.slice(headerStart, cipherStart))) as SecureHeader;
  if (header.schema !== "presentation-studio/encrypted-package" || header.version !== 1 || header.algorithm !== "AES-256-GCM" || header.keyDerivation !== "PBKDF2-SHA-256") {
    throw new Error("The encrypted project format is not supported.");
  }
  if (!Number.isInteger(header.iterations) || header.iterations < 100_000 || header.iterations > 2_000_000) throw new Error("The encrypted project key-derivation settings are unsafe or invalid.");
  const salt = base64ToBytes(header.salt);
  const iv = base64ToBytes(header.iv);
  if (salt.byteLength !== 16 || iv.byteLength !== 12) throw new Error("The encrypted project parameters are invalid.");
  const key = await deriveKey(password, salt, header.iterations);
  try {
    const plaintext = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: toArrayBuffer(iv), additionalData: toArrayBuffer(secureBytes.slice(0, cipherStart)), tagLength: 128 },
      key,
      toArrayBuffer(secureBytes.slice(cipherStart)),
    );
    return new Uint8Array(plaintext);
  } catch {
    throw new Error("The password is incorrect or the encrypted project has been changed.");
  }
}
