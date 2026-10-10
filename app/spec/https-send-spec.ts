import { Account, Contact, Message } from 'mailspring-exports';
import { buildMIME, httpsProviderFor } from '../src/flux/https-send';

const account = (provider: string, smtpUser?: string) =>
  new Account({
    provider,
    emailAddress: 'me@example.com',
    settings: { smtp_username: smtpUser } as any,
  } as any);

describe('https-send', function httpsSend() {
  describe('httpsProviderFor', () => {
    beforeEach(() => {
      spyOn(AppEnv.config, 'get').andCallFake((key) => key === 'core.sending.useHttpsApi');
    });

    it('uses the Gmail API for Gmail accounts', () => {
      expect(httpsProviderFor(account('gmail'))).toBe('gmail');
    });

    it('uses Microsoft Graph for Outlook.com and Microsoft 365 accounts', () => {
      expect(httpsProviderFor(account('outlook', 'me@example.com'))).toBe('graph');
      expect(httpsProviderFor(account('office365'))).toBe('graph');
    });

    it('keeps shared mailboxes on SMTP', () => {
      expect(httpsProviderFor(account('office365', 'owner@example.com'))).toBe(null);
    });

    it('keeps other providers on SMTP', () => {
      expect(httpsProviderFor(account('imap'))).toBe(null);
    });

    it('can be turned off', () => {
      (AppEnv.config.get as any).andReturn(false);
      expect(httpsProviderFor(account('gmail'))).toBe(null);
    });
  });

  describe('buildMIME', () => {
    const draft = (attrs: Partial<Message> = {}) =>
      new Message({
        from: [new Contact({ name: '程诺', email: 'me@example.com' })],
        to: [new Contact({ name: 'To', email: 'to@example.com' })],
        cc: [],
        bcc: [new Contact({ email: 'hidden@example.com' })],
        subject: '[测试] 中文标题',
        headerMessageId: 'ABC@getmailspring.com',
        replyToHeaderMessageId: 'prev@example.com',
        files: [],
        body: '<p>你好</p>',
        ...attrs,
      } as any);

    const headersOf = async (message: Message) => {
      const mime = (await buildMIME(message, message.body)).toString();
      return mime.slice(0, mime.indexOf('\r\n\r\n'));
    };

    it('keeps Bcc so the API delivers to hidden recipients', async () => {
      expect(await headersOf(draft())).toContain('Bcc: hidden@example.com');
    });

    it('uses the draft headerMessageId and threads replies', async () => {
      const headers = await headersOf(draft());
      expect(headers).toContain('Message-ID: <ABC@getmailspring.com>');
      expect(headers).toContain('In-Reply-To: <prev@example.com>');
      expect(headers).toContain('References: <prev@example.com>');
    });

    it('encodes non-ASCII subjects and names', async () => {
      const headers = await headersOf(draft());
      expect(headers).toContain('Subject: =?UTF-8?B?');
      expect(headers).toContain('From: =?UTF-8?B?');
    });

    it('sends plaintext drafts without an HTML part', async () => {
      const message = draft({ plaintext: true, body: 'plain body' } as any);
      const mime = (await buildMIME(message, message.body)).toString();
      expect(mime).toContain('text/plain');
      expect(mime).not.toContain('text/html');
    });
  });
});
