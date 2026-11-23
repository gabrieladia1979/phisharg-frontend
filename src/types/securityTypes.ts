export type AuthResult =
  | "pass"
  | "fail"
  | "softfail"
  | "neutral"
  | "none"
  | "bestguesspass"
  | "unknown";

export interface SecurityFeatures {
  spf_result: AuthResult;
  dkim_result: AuthResult;
  dmarc_result: AuthResult;
  compauth_result: AuthResult;
  return_path: string;
  reply_to: string;
  from_return_path_match: boolean;
  from_reply_to_match: boolean;
  sender_vs_from_match: boolean;
  scl: number;
  threat_category: string;
  spam_filtering_verdict: string;
  auth_as: "Internal" | "Anonymous" | "unknown";
  bcl: number;
  originating_ip: string;
  originating_country: string;
  x_mailer: string;
  received_hop_count: number;
  attachment_count: number;
  attachment_types: string[];
  has_executable_attachment: boolean;
  to_count: number;
  cc_count: number;
  internet_message_id: string;
  has_headers: boolean;
}
