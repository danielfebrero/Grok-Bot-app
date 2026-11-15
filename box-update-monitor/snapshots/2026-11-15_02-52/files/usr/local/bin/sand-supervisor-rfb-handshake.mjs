const RFB_BANNER_LENGTH = 12;
const RFB_SERVER_INIT_HEADER_LENGTH = 24;

export function createRfbHandshake(send, settle) {
  const received = [];
  let receivedLength = 0;
  let stage = "version";
  let securityTypeCount = 0;
  let serverNameLength = 0;
  const take = (length) => {
    if (receivedLength < length) return null;
    const parts = [];
    let offset = 0;
    let consumed = 0;
    while (offset < length) {
      const chunk = received[consumed];
      const size = Math.min(chunk.length, length - offset);
      parts.push(chunk.subarray(0, size));
      offset += size;
      if (size === chunk.length) consumed += 1;
      else received[consumed] = chunk.subarray(size);
    }
    received.splice(0, consumed);
    receivedLength -= length;
    return Buffer.concat(parts, length);
  };
  return (chunk) => {
    if (chunk.length > 0) received.push(chunk);
    receivedLength += chunk.length;
    for (;;) {
      if (stage === "version") {
        const banner = take(RFB_BANNER_LENGTH);
        if (banner == null) return;
        if (banner.toString("ascii") !== "RFB 003.008\n") return settle(false);
        send(Buffer.from("RFB 003.008\n"));
        stage = "security-count";
      } else if (stage === "security-count") {
        const count = take(1);
        if (count == null) return;
        securityTypeCount = count[0];
        if (securityTypeCount === 0) return settle(false);
        stage = "security-types";
      } else if (stage === "security-types") {
        const types = take(securityTypeCount);
        if (types == null) return;
        if (!types.includes(1)) return settle(false);
        send(Buffer.from([1]));
        stage = "security-result";
      } else if (stage === "security-result") {
        const result = take(4);
        if (result == null) return;
        if (result.readUInt32BE(0) !== 0) return settle(false);
        send(Buffer.from([1]));
        stage = "server-init";
      } else if (stage === "server-init") {
        const header = take(RFB_SERVER_INIT_HEADER_LENGTH);
        if (header == null) return;
        const width = header.readUInt16BE(0);
        const height = header.readUInt16BE(2);
        const bitsPerPixel = header[4];
        const depth = header[5];
        serverNameLength = header.readUInt32BE(20);
        if (
          width === 0 ||
          height === 0 ||
          ![8, 16, 32].includes(bitsPerPixel) ||
          depth === 0 ||
          depth > bitsPerPixel ||
          serverNameLength > 64 * 1024
        ) {
          return settle(false);
        }
        stage = "server-name";
      } else {
        const name = take(serverNameLength);
        if (name == null) return;
        return settle(true);
      }
    }
  };
}

export function parseForkRfbTokenRecord(fileName, raw) {
  if (!/^[0-9]+$/.test(fileName)) return null;
  const display = Number.parseInt(fileName, 10);
  if (!Number.isInteger(display) || display < 2) return null;
  const match = /^([^:\s]+):\s+(.+):([0-9]+)\s*$/.exec(raw);
  if (match == null || match[1] !== fileName) return null;
  const vncPort = Number.parseInt(match[3], 10);
  if (!Number.isInteger(vncPort) || vncPort <= 0 || vncPort > 65_535) {
    return null;
  }
  return { display, token: fileName, vncPort };
}
