import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import { ICurrentUser, IEmailNotification } from './models';

export class GraphService {
  private _graphClientFactory: MSGraphClientFactory;

  constructor(graphClientFactory: MSGraphClientFactory) {
    this._graphClientFactory = graphClientFactory;
  }

  private async _getClient(): Promise<MSGraphClientV3> {
    return this._graphClientFactory.getClient('3');
  }

  // ── Current user ───────────────────────────────────────────────────────────

  public async getCurrentUser(): Promise<ICurrentUser> {
    try {
      const client = await this._getClient();
      const response = await client
        .api('/me')
        .select('displayName,mail,userPrincipalName')
        .get() as { displayName: string; mail: string; userPrincipalName: string };

      return {
        displayName: response.displayName ?? 'Unknown User',
        email: response.mail ?? response.userPrincipalName ?? '',
      };
    } catch (err) {
      throw new Error(`Failed to get current user: ${(err as Error).message}`);
    }
  }

  // ── Send email via Microsoft Graph Mail.Send ───────────────────────────────

  public async sendEmail(notification: IEmailNotification): Promise<void> {
    try {
      const client = await this._getClient();

      const toRecipients = notification.to.map(email => ({
        emailAddress: { address: email },
      }));

      const ccRecipients = (notification.cc ?? []).map(email => ({
        emailAddress: { address: email },
      }));

      const message: Record<string, unknown> = {
        subject: notification.subject,
        body: {
          contentType: 'HTML',
          content: notification.bodyHtml,
        },
        toRecipients,
      };

      if (ccRecipients.length > 0) {
        message.ccRecipients = ccRecipients;
      }

      await client.api('/me/sendMail').post({ message, saveToSentItems: true });
    } catch (err) {
      throw new Error(`Failed to send email: ${(err as Error).message}`);
    }
  }

  // ── Convenience: get user profile photo URL ────────────────────────────────

  public async getUserPhotoUrl(email: string): Promise<string> {
    try {
      const client = await this._getClient();
      // Returns the photo as a base64 data URL
      const response = await client
        .api(`/users/${email}/photo/$value`)
        .get() as ArrayBuffer;

      const base64 = btoa(
        new Uint8Array(response).reduce((data, byte) => data + String.fromCharCode(byte), '')
      );
      return `data:image/jpeg;base64,${base64}`;
    } catch {
      return '';
    }
  }
}
