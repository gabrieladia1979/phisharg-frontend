import { SecurityFeatures, AuthResult } from "../types/securityTypes";

const DANGEROUS_EXTENSIONS = new Set([
  ".exe", ".bat", ".cmd", ".com", ".vbs", ".vbe",
  ".js", ".jse", ".wsf", ".wsh", ".msi", ".scr",
  ".pif", ".ps1", ".hta", ".cpl", ".reg", ".inf",
]);

function getInternetHeaders(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve) => {
    if (
      !Office.context.requirements.isSetSupported("Mailbox", "1.8") ||
      !item.getAllInternetHeadersAsync
    ) {
      resolve("");
      return;
    }

    item.getAllInternetHeadersAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }
      resolve("");
    });
  });
}

function extractDomain(email: string): string {
  const atIndex = email.lastIndexOf("@");
  if (atIndex === -1) {
    return "";
  }
  return email.substring(atIndex + 1).toLowerCase().trim();
}

function extractEmailFromHeader(value: string): string {
  const angleMatch = value.match(/<([^>]+)>/);
  if (angleMatch) {
    return angleMatch[1].trim().toLowerCase();
  }
  const emailMatch = value.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/);
  if (emailMatch) {
    return emailMatch[0].toLowerCase();
  }
  return value.trim().toLowerCase();
}

function parseAuthenticationResults(headers: string): {
  spf: AuthResult;
  dkim: AuthResult;
  dmarc: AuthResult;
  compauth: AuthResult;
} {
  const result = {
    spf: "unknown" as AuthResult,
    dkim: "unknown" as AuthResult,
    dmarc: "unknown" as AuthResult,
    compauth: "unknown" as AuthResult,
  };

  const spfMatch = headers.match(/spf=(pass|fail|softfail|neutral|none)/i);
  if (spfMatch) {
    result.spf = spfMatch[1].toLowerCase() as AuthResult;
  }

  const dkimMatch = headers.match(/dkim=(pass|fail|neutral|none)/i);
  if (dkimMatch) {
    result.dkim = dkimMatch[1].toLowerCase() as AuthResult;
  }

  const dmarcMatch = headers.match(
    /dmarc=(pass|fail|bestguesspass|none)/i
  );
  if (dmarcMatch) {
    result.dmarc = dmarcMatch[1].toLowerCase() as AuthResult;
  }

  const compMatch = headers.match(/compauth=(pass|fail)/i);
  if (compMatch) {
    result.compauth = compMatch[1].toLowerCase() as AuthResult;
  }

  return result;
}

function parseAntispamHeaders(headers: string): {
  scl: number;
  threatCategory: string;
  sfv: string;
  bcl: number;
  originatingCountry: string;
  connectingIp: string;
} {
  const result = {
    scl: -1,
    threatCategory: "NONE",
    sfv: "",
    bcl: -1,
    originatingCountry: "",
    connectingIp: "",
  };

  const sclMatch = headers.match(/SCL:(\-?\d+)/i);
  if (sclMatch) {
    result.scl = parseInt(sclMatch[1], 10);
  }

  const catMatch = headers.match(/CAT:([A-Z]+)/i);
  if (catMatch) {
    result.threatCategory = catMatch[1].toUpperCase();
  }

  const sfvMatch = headers.match(/SFV:([A-Z]+)/i);
  if (sfvMatch) {
    result.sfv = sfvMatch[1].toUpperCase();
  }

  const bclMatch = headers.match(/BCL:(\d+)/i);
  if (bclMatch) {
    result.bcl = parseInt(bclMatch[1], 10);
  }

  const ctryMatch = headers.match(/CTRY:([A-Z]{2})/i);
  if (ctryMatch) {
    result.originatingCountry = ctryMatch[1].toUpperCase();
  }

  const cipMatch = headers.match(/CIP:([\d.]+)/i);
  if (cipMatch) {
    result.connectingIp = cipMatch[1];
  }

  return result;
}

function parseSpoofingIndicators(
  headers: string,
  fromEmail: string
): {
  returnPath: string;
  replyTo: string;
  fromReturnPathMatch: boolean;
  fromReplyToMatch: boolean;
} {
  const fromDomain = extractDomain(fromEmail);

  const returnPathMatch = headers.match(/Return-Path:\s*<?([^>\r\n]+)>?/im);
  const returnPath = returnPathMatch
    ? extractEmailFromHeader(returnPathMatch[1])
    : "";
  const returnPathDomain = extractDomain(returnPath);

  const replyToMatch = headers.match(/Reply-To:\s*<?([^>\r\n]+)>?/im);
  const replyTo = replyToMatch
    ? extractEmailFromHeader(replyToMatch[1])
    : "";
  const replyToDomain = extractDomain(replyTo);

  return {
    returnPath,
    replyTo,
    fromReturnPathMatch:
      returnPath === "" || fromDomain === returnPathDomain,
    fromReplyToMatch:
      replyTo === "" || fromDomain === replyToDomain,
  };
}

function parseOriginMetadata(headers: string): {
  originatingIp: string;
  xMailer: string;
  receivedHopCount: number;
} {
  const origIpMatch = headers.match(
    /X-Originating-IP:\s*\[?([^\]\r\n]+)\]?/im
  );
  const originatingIp = origIpMatch ? origIpMatch[1].trim() : "";

  const mailerMatch = headers.match(/(?:X-Mailer|User-Agent):\s*([^\r\n]+)/im);
  const xMailer = mailerMatch ? mailerMatch[1].trim() : "";

  const receivedMatches = headers.match(/^Received:/gim);
  const receivedHopCount = receivedMatches ? receivedMatches.length : 0;

  return { originatingIp, xMailer, receivedHopCount };
}

function parseAuthAs(headers: string): "Internal" | "Anonymous" | "unknown" {
  const match = headers.match(
    /X-MS-Exchange-Organization-AuthAs:\s*(Internal|Anonymous)/im
  );
  if (match) {
    return match[1] as "Internal" | "Anonymous";
  }
  return "unknown";
}

function getAttachmentFeatures(item: Office.MessageRead): {
  count: number;
  types: string[];
  hasExecutable: boolean;
} {
  const attachments = item.attachments || [];
  const types: string[] = [];
  let hasExecutable = false;

  for (const att of attachments) {
    if (att.contentType) {
      types.push(att.contentType);
    }

    const name = (att.name || "").toLowerCase();
    const dotIndex = name.lastIndexOf(".");
    if (dotIndex !== -1) {
      const ext = name.substring(dotIndex);
      if (DANGEROUS_EXTENSIONS.has(ext)) {
        hasExecutable = true;
      }
    }
  }

  return {
    count: attachments.length,
    types,
    hasExecutable,
  };
}

function createEmptyFeatures(): SecurityFeatures {
  return {
    spf_result: "unknown",
    dkim_result: "unknown",
    dmarc_result: "unknown",
    compauth_result: "unknown",
    return_path: "",
    reply_to: "",
    from_return_path_match: true,
    from_reply_to_match: true,
    sender_vs_from_match: true,
    scl: -1,
    threat_category: "NONE",
    spam_filtering_verdict: "",
    auth_as: "unknown",
    bcl: -1,
    originating_ip: "",
    originating_country: "",
    x_mailer: "",
    received_hop_count: 0,
    attachment_count: 0,
    attachment_types: [],
    has_executable_attachment: false,
    to_count: 0,
    cc_count: 0,
    internet_message_id: "",
    has_headers: false,
  };
}

export async function extractSecurityFeatures(
  item: Office.MessageRead
): Promise<SecurityFeatures> {
  const features = createEmptyFeatures();

  const fromEmail = item.from?.emailAddress?.toLowerCase() || "";
  const senderEmail = item.sender?.emailAddress?.toLowerCase() || "";
  features.sender_vs_from_match =
    fromEmail === senderEmail || senderEmail === "";

  features.internet_message_id = item.internetMessageId || "";

  const attachInfo = getAttachmentFeatures(item);
  features.attachment_count = attachInfo.count;
  features.attachment_types = attachInfo.types;
  features.has_executable_attachment = attachInfo.hasExecutable;

  features.to_count = (item.to || []).length;
  features.cc_count = (item.cc || []).length;

  const headers = await getInternetHeaders(item);

  if (!headers) {
    features.has_headers = false;
    return features;
  }

  features.has_headers = true;

  const auth = parseAuthenticationResults(headers);
  features.spf_result = auth.spf;
  features.dkim_result = auth.dkim;
  features.dmarc_result = auth.dmarc;
  features.compauth_result = auth.compauth;

  const antispam = parseAntispamHeaders(headers);
  features.scl = antispam.scl;
  features.threat_category = antispam.threatCategory;
  features.spam_filtering_verdict = antispam.sfv;
  features.bcl = antispam.bcl;
  features.originating_country = antispam.originatingCountry;

  features.auth_as = parseAuthAs(headers);

  const spoofing = parseSpoofingIndicators(headers, fromEmail);
  features.return_path = spoofing.returnPath;
  features.reply_to = spoofing.replyTo;
  features.from_return_path_match = spoofing.fromReturnPathMatch;
  features.from_reply_to_match = spoofing.fromReplyToMatch;

  const origin = parseOriginMetadata(headers);
  features.originating_ip =
    origin.originatingIp || antispam.connectingIp;
  features.x_mailer = origin.xMailer;
  features.received_hop_count = origin.receivedHopCount;

  return features;
}
