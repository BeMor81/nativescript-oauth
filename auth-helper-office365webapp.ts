/// <reference path="references.d.ts" />

import * as tnsOauth from './tns-oauth';
import { AuthHelper } from './auth-helper';
import * as TnsOAuth from './tns-oauth-interfaces';

export class AuthHelperOffice365WebApp extends AuthHelper implements TnsOAuth.ITnsAuthHelper {

  constructor(tenant: string, clientId: string, clientSecret: string, scope: Array<string>, loginHint?: string) {
    super();
    var scopeStr = scope.join('%20');
    this.credentials = {
      authority: 'https://login.microsoftonline.com/' + tenant,
      authorizeEndpoint: '/oauth2/v2.0/authorize',
      tokenEndpoint: '/oauth2/v2.0/token',
      clientId: clientId,
      clientSecret: clientSecret,
      redirectUri: 'urn:ietf:wg:oauth:2.0:oob',
      scope: scopeStr,
      loginHint: loginHint
    };
  }

  public logout(successPage?: string): Promise<void> {
    let cookieDomains = ["login.microsoftonline.com", ".live.com"];
    return this._logout(successPage, cookieDomains);
  }
}