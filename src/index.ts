import { Transport } from './utils/transport';
import {
  IEmailGetResponse,
  IEmailQueryResponse,
  IEmailSetResponse,
  IArguments,
  IMailboxGetResponse,
  IMailboxSetResponse,
  ISession,
  IEmailGetArguments,
  IMailboxGetArguments,
  IMailboxSetArguments,
  IMethodName,
  IReplaceableAccountId,
  IEmailQueryArguments,
  IEmailSetArguments,
  IMailboxChangesArguments,
  IMailboxChangesResponse,
  IEmailSubmissionSetArguments,
  IEmailSubmissionGetResponse,
  IEmailSubmissionGetArguments,
  IEmailSubmissionChangesArguments,
  IEmailSubmissionSetResponse,
  IEmailSubmissionChangesResponse,
  IEmailChangesArguments,
  IEmailChangesResponse,
  IInvocation,
  IUploadResponse,
  IEmailImportArguments,
  IEmailImportResponse,
  IThreadGetArguments,
  IThreadGetResponse,
  IRequest,
} from './types';

export class Client {
  private readonly DEFAULT_USING = ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'];

  private transport: Transport;
  private httpHeaders: { [headerName: string]: string };

  private sessionUrl: string;
  private overriddenApiUrl?: string;
  private session?: ISession;
  protected accessTokenProvider: string | Promise<string> | (() => string) | (() => Promise<String>);

  constructor({
    sessionUrl,
    accessTokenProvider,
    overriddenApiUrl,
    transport,
    httpHeaders,
  }: {
    sessionUrl: string;
    accessTokenProvider: string | Promise<string> | (() => string) | (() => Promise<String>);
    overriddenApiUrl?: string;
    transport: Transport;
    httpHeaders?: { [headerName: string]: string };
  }) {
    this.sessionUrl = sessionUrl;
    if (overriddenApiUrl) {
      this.overriddenApiUrl = overriddenApiUrl;
    }
    this.transport = transport;
    this.accessTokenProvider = accessTokenProvider;
    this.httpHeaders = {
      Accept: 'application/json;jmapVersion=rfc-8621',
      // Authorization: `Bearer ${accessToken}`,
      ...(httpHeaders ? httpHeaders : {}),
    };
  }

  protected async getAccessToken() {
    switch (typeof this.accessTokenProvider) {
      case 'string':
        return this.accessTokenProvider;
      case 'object':
        if (this.accessTokenProvider instanceof Promise) {
          const accessToken = await this.accessTokenProvider;
          if (typeof accessToken === 'string') {
            return accessToken;
          }
        }
        break;
      case 'function':
        const accessToken = await this.accessTokenProvider();
        if (typeof accessToken === 'string') {
          return accessToken;
        }
      default:
        throw new Error('Access Token Provider not resolving to a string')
    }
  }

  public async fetchSession(sessionHeaders?: { [headerName: string]: string }): Promise<void> {
    const accessToken = await this.getAccessToken();
    const requestHeaders = {
      ...this.httpHeaders,
      Authorization: `Bearer ${accessToken}`,
      ...(sessionHeaders ? sessionHeaders : {}),

    };
    const sessionPromise = this.transport.get<ISession>(this.sessionUrl, requestHeaders);
    return sessionPromise.then(session => {
      this.session = session;
      return;
    });
  }

  public getSession(): ISession {
    if (!this.session) {
      throw new Error('Undefined session, should call fetchSession and wait for its resolution');
    }
    return this.session;
  }

  public getAccountIds(): string[] {
    const session = this.getSession();

    return Object.keys(session.accounts);
  }

  public getFirstAccountId(): string {
    const accountIds = this.getAccountIds();

    if (accountIds.length === 0) {
      throw new Error('No account available for this session');
    }

    return accountIds[0];
  }

  public mailbox_get(args: IMailboxGetArguments): Promise<IMailboxGetResponse> {
    return this.request<IMailboxGetResponse>('Mailbox/get', args);
  }

  public mailbox_changes(args: IMailboxChangesArguments): Promise<IMailboxChangesResponse> {
    return this.request<IMailboxChangesResponse>('Mailbox/changes', args);
  }

  public mailbox_set(args: IMailboxSetArguments): Promise<IMailboxSetResponse> {
    return this.request<IMailboxSetResponse>('Mailbox/set', args);
  }

  public email_get(args: IEmailGetArguments): Promise<IEmailGetResponse> {
    return this.request<IEmailGetResponse>('Email/get', args);
  }

  public email_changes(args: IEmailChangesArguments): Promise<IEmailChangesResponse> {
    return this.request<IEmailChangesResponse>('Email/changes', args);
  }

  public email_query(args: IEmailQueryArguments): Promise<IEmailQueryResponse> {
    return this.request<IEmailQueryResponse>('Email/query', args);
  }

  public email_set(args: IEmailSetArguments): Promise<IEmailSetResponse> {
    return this.request<IEmailSetResponse>('Email/set', args);
  }

  public email_import(args: IEmailImportArguments): Promise<IEmailImportResponse> {
    return this.request<IEmailImportResponse>('Email/import', args);
  }

  public thread_get(args: IThreadGetArguments): Promise<IThreadGetResponse> {
    return this.request<IThreadGetResponse>('Thread/get', args);
  }

  public emailSubmission_get(
    args: IEmailSubmissionGetArguments,
  ): Promise<IEmailSubmissionGetResponse> {
    return this.request<IEmailSubmissionGetResponse>('EmailSubmission/get', args);
  }

  public emailSubmission_changes(
    args: IEmailSubmissionChangesArguments,
  ): Promise<IEmailSubmissionChangesResponse> {
    return this.request<IEmailSubmissionChangesResponse>('EmailSubmission/changes', args);
  }

  public emailSubmission_set(
    args: IEmailSubmissionSetArguments,
  ): Promise<IEmailSubmissionSetResponse> {
    return this.request<IEmailSubmissionSetResponse>('EmailSubmission/set', args);
  }

  public upload(buffer: ArrayBuffer, type = 'application/octet-stream'): Promise<IUploadResponse> {
    const uploadUrl = this.getSession().uploadUrl;
    const accountId = this.getFirstAccountId();
    const requestHeaders = {
      ...this.httpHeaders,
      'Content-Type': type,
    };
    return this.transport.post<IUploadResponse>(
      uploadUrl.replace('{accountId}', encodeURIComponent(accountId)),
      buffer,
      requestHeaders,
    );
  }

  private async request<ResponseType>(methodName: IMethodName, args: IArguments) {
    return this.rawRequest<ResponseType>([[methodName, this.replaceAccountId(args), '0']])
      .then(response => {
        const methodResponse = response.methodResponses[0];

        if (methodResponse[0] === 'error') {
          throw methodResponse[1];
        }

        return methodResponse[1];
      });;
  }

  public async rawRequest<R>(methodCalls: IRequest['methodCalls']) {
    const apiUrl = this.overriddenApiUrl || this.getSession().apiUrl;
    const accessToken = await this.getAccessToken();
    return await this.transport
      .post<{
        sessionState: string;
        methodResponses: IInvocation<R>[];
      }>(
        apiUrl,
        {
          using: this.getCapabilities(),
          methodCalls,
        },
        {
          ...this.httpHeaders,
          Authorization: `Bearer ${accessToken}`,
        },
      );
  }

  private replaceAccountId<U extends IReplaceableAccountId>(input: U): U {
    return input.accountId !== null
      ? input
      : {
        ...input,
        accountId: this.getFirstAccountId(),
      };
  }

  private getCapabilities() {
    return this.session?.capabilities ? Object.keys(this.session.capabilities) : this.DEFAULT_USING;
  }
}
