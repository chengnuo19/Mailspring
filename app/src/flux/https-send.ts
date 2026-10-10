import { localized } from '../intl';
import KeyManager from '../key-manager';
import { Account } from './models/account';
import { Contact } from './models/contact';
import { Message } from './models/message';
import AttachmentStore from './stores/attachment-store';
import {
  GMAIL_CLIENT_ID,
  GMAIL_CLIENT_SECRET,
  O365_CLIENT_ID,
} from '../../internal_packages/onboarding/lib/onboarding-constants';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const MailComposerModule = require('nodemailer/lib/mail-composer');
const MailComposer = MailComposerModule.default || MailComposerModule;

/*
Sends drafts for Microsoft 365 / Outlook.com and Gmail accounts through the providers' HTTPS
APIs (Microsoft Graph `sendMail`, Gmail API `messages.send`) instead of SMTP, as the new
Outlook for Windows does. Networks that block outbound SMTP ports (465/587) — DigitalOcean
droplets, many campus networks — can still send, and the provider files the message in Sent.
Receiving stays on IMAP in the sync engine.
*/

type HttpsProvider = 'graph' | 'gmail';

const GRAPH_SEND_SCOPE = 'https://graph.microsoft.com/Mail.Send offline_access';
// Graph rejects request bodies over 4 MB, and the MIME body is sent base64-encoded.
const GRAPH_MAX_MIME_BYTES = Math.floor((4 * 1024 * 1024 * 3) / 4) - 64 * 1024;
const REQUEST_TIMEOUT_MS = 120 * 1000;

/** Thrown when this account cannot use HTTPS sending; the caller falls back to SMTP. */
export class HttpsSendUnavailableError extends Error {}

export function httpsProviderFor(account: Account): HttpsProvider | null {
  if (!AppEnv.config.get('core.sending.useHttpsApi')) {
    return null;
  }
  if (account.provider === 'gmail') {
    return 'gmail';
  }
  // Graph's /me/sendMail sends as the signed-in user, so shared mailboxes (whose SMTP login
  // is the owner, not the mailbox address) keep using SMTP "Send As".
  const smtpUser = account.settings.smtp_username;
  if (
    (account.provider === 'office365' || account.provider === 'outlook') &&
    (!smtpUser || smtpUser.toLowerCase() === account.emailAddress.toLowerCase())
  ) {
    return 'graph';
  }
  return null;
}

const tokenCache = new Map<string, { token: string; expiresAt: number }>();

async function fetchAccessToken(account: Account, provider: HttpsProvider) {
  const cached = tokenCache.get(account.id);
  if (cached && cached.expiresAt > Date.now() + 60 * 1000) {
    return cached.token;
  }

  const { settings } = await KeyManager.insertAccountSecrets(account);
  const refreshClientId = (settings as any).refresh_client_id as string | undefined;
  if (!settings.refresh_token) {
    throw new HttpsSendUnavailableError('The account has no OAuth refresh token.');
  }

  let url: string;
  let form: Record<string, string>;
  if (provider === 'gmail') {
    // Older Gmail accounts were authorized through a client whose secret this app lacks.
    if (refreshClientId && refreshClientId !== GMAIL_CLIENT_ID) {
      throw new HttpsSendUnavailableError(`Unsupported Gmail OAuth client ${refreshClientId}.`);
    }
    url = 'https://oauth2.googleapis.com/token';
    form = {
      client_id: GMAIL_CLIENT_ID,
      client_secret: GMAIL_CLIENT_SECRET,
      refresh_token: settings.refresh_token,
      grant_type: 'refresh_token',
    };
  } else {
    url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    form = {
      client_id: refreshClientId || O365_CLIENT_ID,
      refresh_token: settings.refresh_token,
      grant_type: 'refresh_token',
      scope: GRAPH_SEND_SCOPE,
    };
  }

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
    body: new URLSearchParams(form).toString(),
    signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
  });
  const json = await resp.json().catch(() => ({}));
  if (!resp.ok || !json.access_token) {
    // invalid_grant / invalid_scope / consent_required: the grant predates Mail.Send consent
    // (Microsoft accounts added before it was requested) or was revoked.
    throw new HttpsSendUnavailableError(
      `Token request failed (${resp.status}): ${json.error || ''} ${json.error_description || ''}`
    );
  }
  tokenCache.set(account.id, {
    token: json.access_token,
    expiresAt: Date.now() + (Number(json.expires_in) || 3600) * 1000,
  });
  return json.access_token as string;
}

const toAddress = (c: Contact) => ({ name: c.name || '', address: c.email });
const bracketed = (id: string) => (id.startsWith('<') ? id : `<${id}>`);

function plainTextFromHTML(html: string) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return (doc.body.textContent || '').trim();
}

export function buildMIME(draft: Message, body: string): Promise<Buffer> {
  const mail: Record<string, any> = {
    from: toAddress(draft.from[0]),
    to: draft.to.map(toAddress),
    cc: draft.cc.map(toAddress),
    bcc: draft.bcc.map(toAddress),
    replyTo: draft.replyTo.length ? draft.replyTo.map(toAddress) : undefined,
    subject: draft.subject || '',
    messageId: bracketed(draft.headerMessageId),
    date: new Date(),
    attachments: draft.files.map((file) => ({
      filename: file.filename,
      path: AttachmentStore.pathForFile(file),
      contentType: file.contentType || undefined,
      cid: file.contentId || undefined,
    })),
  };
  if (draft.replyToHeaderMessageId) {
    mail.inReplyTo = bracketed(draft.replyToHeaderMessageId);
    mail.references = bracketed(draft.replyToHeaderMessageId);
  }
  if (draft.plaintext) {
    mail.text = body;
  } else {
    mail.html = body;
    mail.text = plainTextFromHTML(body);
  }

  const node = new MailComposer(mail).compile();
  // The APIs read recipients from the headers, so Bcc must stay in the message.
  node.keepBcc = true;
  return new Promise((resolve, reject) => {
    node.build((err: Error, message: Buffer) => (err ? reject(err) : resolve(message)));
  });
}

async function apiError(resp: Response) {
  const text = await resp.text();
  try {
    const json = JSON.parse(text);
    const message = json.error?.message || json.error_description || text;
    return `${resp.status} ${message}`;
  } catch {
    return `${resp.status} ${text.slice(0, 300)}`;
  }
}

async function postMIME(provider: HttpsProvider, token: string, mime: Buffer) {
  if (provider === 'graph') {
    return fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'text/plain' },
      body: mime.toString('base64'),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
  }
  // The media upload endpoint accepts messages up to 35 MB.
  return fetch(
    'https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media',
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'message/rfc822' },
      body: new Uint8Array(mime),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    }
  );
}

/**
 * Sends `draft` with `body` (the body after composer extension transforms). Rejects with
 * HttpsSendUnavailableError when SMTP should be used instead, or a user-facing Error.
 */
export async function sendOverHttps(account: Account, draft: Message, body: string) {
  const provider = httpsProviderFor(account);
  if (!provider) {
    throw new HttpsSendUnavailableError('HTTPS sending is not available for this account.');
  }
  const mime = await buildMIME(draft, body);
  if (provider === 'graph' && mime.length > GRAPH_MAX_MIME_BYTES) {
    throw new Error(
      localized(
        'This message is too large to send through Microsoft Graph (about 3 MB with attachments). Remove some attachments and try again.'
      )
    );
  }

  let token = await fetchAccessToken(account, provider);
  let resp = await postMIME(provider, token, mime);
  if (resp.status === 401) {
    tokenCache.delete(account.id);
    token = await fetchAccessToken(account, provider);
    resp = await postMIME(provider, token, mime);
  }
  if (provider === 'gmail' && resp.status === 403) {
    const detail = await apiError(resp);
    // The Gmail API is not enabled for the OAuth client's Google Cloud project.
    if (/SERVICE_DISABLED|accessNotConfigured|has not been used/i.test(detail)) {
      throw new HttpsSendUnavailableError(detail);
    }
    throw new Error(detail);
  }
  if (!resp.ok) {
    throw new Error(await apiError(resp));
  }
}
