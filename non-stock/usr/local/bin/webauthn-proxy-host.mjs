import { Buffer } from "node:buffer";
import { readFileSync } from "node:fs";

import { isWebAuthnArmed, SAND_WEBAUTHN_STALE_TAB_MESSAGE } from "./webauthn-arm.mjs";

// Chrome starts a new native messaging host process for each sendNativeMessage, frames every message on stdio as a 32-bit native-byte-order (little-endian here) length followed by that many bytes of UTF-8 JSON, and takes the host's first message as the response: https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging#native-messaging-host-protocol
const HEADER_BYTES = 4;
// Chrome caps a message sent to a native messaging host at 64 MiB: https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging#native-messaging-host-protocol
const MAX_MESSAGE_BYTES = 64 * 1024 * 1024;

// >>> box port table (from sand/src/shared/box/box-contract.ts; regenerate: pnpm --filter sand run gen:box-ports) >>>
const SAND_BOX_PORT_HOST_GATEWAY = "1340";
// <<< box port table <<<

function readEnv(name, fallback) {
  const value = process.env[name];
  return value === undefined || value === "" ? fallback : value;
}

function credentialFile() {
  const primaryDir = readEnv("SAND_PRIMARY_XDG_RUNTIME_DIR", "/tmp/xdg-runtime-box");
  for (const runtimeDir of [readEnv("XDG_RUNTIME_DIR", undefined), primaryDir]) {
    if (runtimeDir === undefined) {
      continue;
    }
    try {
      const [token, port] = readFileSync(`${runtimeDir}/sand-gateway-credential`, "utf8").split(
        "\n",
      );
      if (token !== undefined && token !== "") {
        return { token, port };
      }
    } catch (error) {
      if (error?.code !== "ENOENT") {
        console.error(
          `webauthn-proxy-host: credential file in ${runtimeDir} unreadable: ${String(error)}`,
        );
      }
    }
  }
  return undefined;
}

function gatewayBaseUrl(credential) {
  const port =
    readEnv("SAND_HOST_PORT", undefined) ?? credential?.port ?? SAND_BOX_PORT_HOST_GATEWAY;
  return `http://127.0.0.1:${port}`;
}

function writeMessage(payload) {
  const body = Buffer.from(JSON.stringify(payload), "utf8");
  const header = Buffer.alloc(HEADER_BYTES);
  header.writeUInt32LE(body.length, 0);
  process.stdout.write(Buffer.concat([header, body]));
}

function failure(name, message) {
  return { ok: false, error: { name, message } };
}

async function readMessage() {
  const chunks = [];
  let total = 0;
  for await (const chunk of process.stdin) {
    chunks.push(chunk);
    total += chunk.length;
    if (total > MAX_MESSAGE_BYTES) {
      throw new Error("native message exceeded the maximum size");
    }
    const buffered = Buffer.concat(chunks, total);
    if (buffered.length < HEADER_BYTES) {
      continue;
    }
    const length = buffered.readUInt32LE(0);
    if (buffered.length >= HEADER_BYTES + length) {
      return JSON.parse(buffered.subarray(HEADER_BYTES, HEADER_BYTES + length).toString("utf8"));
    }
  }
  return undefined;
}

function webAuthnArmPaths() {
  return {
    armDir: readEnv("SAND_WEBAUTHN_ARM_DIR", undefined),
    armUntilPath: readEnv("SAND_WEBAUTHN_ARM_UNTIL_PATH", undefined),
    handoffArmUntilPath: readEnv("SAND_WEBAUTHN_HANDOFF_ARM_UNTIL_PATH", undefined),
  };
}

async function requestCeremony(message) {
  if (!isWebAuthnArmed(Date.now(), webAuthnArmPaths())) {
    return failure("NotAllowedError", SAND_WEBAUTHN_STALE_TAB_MESSAGE);
  }

  const credential = credentialFile();
  const token = readEnv("SAND_GATEWAY_TOKEN", undefined) ?? credential?.token;
  if (token === undefined) {
    return failure(
      "NotAllowedError",
      "Sand's in-box gateway token is not available to the browser bridge.",
    );
  }

  const response = await fetch(`${gatewayBaseUrl(credential)}/api/requestWebAuthnCeremony`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${token}`,
    },
    body: JSON.stringify({
      kind: message.kind,
      origin: message.origin,
      optionsJson: message.optionsJson,
    }),
  });
  if (!response.ok) {
    return failure(
      "NotAllowedError",
      `Sand's in-box host refused the ceremony (HTTP ${response.status}).`,
    );
  }
  return await response.json();
}

async function main() {
  let message;
  try {
    message = await readMessage();
  } catch (error) {
    writeMessage(failure("DataError", `unreadable native message: ${error}`));
    return;
  }
  if (message === undefined) {
    writeMessage(failure("DataError", "no native message was received"));
    return;
  }

  try {
    writeMessage(await requestCeremony(message));
  } catch (error) {
    writeMessage(
      failure("NotAllowedError", `could not reach Sand's in-box host: ${error?.message ?? error}`),
    );
  }
}

await main();
