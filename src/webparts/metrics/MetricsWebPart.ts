import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import { SPHttpClient, HttpClient, HttpClientResponse, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';

import * as strings from 'MetricsWebPartStrings';
import { MetricsConfig } from '../../core/types';
import { parseConfig } from '../../core/validate';
import { DataContext, IHttp } from '../../core/datasources';
import { MetricsWebPartView } from './MetricsWebPartView';
import starter from '../../samples/starter.json';

export interface IMetricsWebPartProps {
  configSource: 'inline' | 'url';
  configJson: string;
  configUrl: string;
}

export default class MetricsWebPart extends BaseClientSideWebPart<IMetricsWebPartProps> {
  private _root: Root | undefined;
  private _themeProvider!: ThemeProvider;
  private _theme: IReadonlyTheme | undefined;
  private _data!: DataContext;
  private _config: MetricsConfig | undefined;
  private _configError: string | undefined;
  private _reloadToken: number = 0;
  private _editorOpen: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._onThemeChanged);

    const http: IHttp = {
      getJson: async (url, options) => {
        const request: IHttpClientOptions = { headers: options.headers };
        const response = options.sharepoint
          ? await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, request)
          : await this.context.httpClient.get(url, HttpClient.configurations.v1, request);
        if (!response.ok) throw new Error(await errorMessage(response));
        return response.json();
      },
      postJson: async (url, options) => {
        const request: ISPHttpClientOptions = {
          headers: options.headers,
          body: options.body === undefined ? undefined : JSON.stringify(options.body),
        };
        const response = options.sharepoint
          ? await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request)
          : await this.context.httpClient.post(url, HttpClient.configurations.v1, request);
        if (!response.ok) throw new Error(await errorMessage(response));
        const text = await response.text();
        return text ? JSON.parse(text) : undefined;
      },
      getBinary: async (url, options) => {
        const request: IHttpClientOptions = { headers: options.headers };
        const response = options.sharepoint
          ? await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, request)
          : await this.context.httpClient.get(url, HttpClient.configurations.v1, request);
        if (!response.ok) throw new Error(await errorMessage(response));
        return response.arrayBuffer();
      },
    };
    const user = this.context.pageContext.user;
    this._data = {
      http,
      webUrl: this.context.pageContext.web.absoluteUrl,
      user: { displayName: user.displayName, email: user.email, loginName: user.loginName },
    };
    await this._loadConfig();
  }

  public render(): void {
    if (!this._root) this._root = createRoot(this.domElement);
    this._root.render(
      React.createElement(MetricsWebPartView, {
        config: this._config,
        error: this._configError,
        data: this._data,
        sectionTheme: this._theme,
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
        reloadToken: this._reloadToken,
        instanceKey: this.instanceId.slice(0, 8),
        isEditMode: this.displayMode === DisplayMode.Edit,
        editorOpen: this._editorOpen,
        editorText: this.properties.configJson || JSON.stringify(starter, undefined, 2),
        onOpenEditor: () => this._setEditorOpen(true),
        onCloseEditor: () => this._setEditorOpen(false),
        onSaveConfig: (text: string) => this._saveConfig(text),
        onUseStarter: () => this._saveConfig(JSON.stringify(starter, undefined, 2)),
      }),
    );
  }

  protected onDisplayModeChanged(): void {
    this.render();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (oldValue === newValue) return;
    if (propertyPath === 'configJson' || propertyPath === 'configSource' || propertyPath === 'configUrl') {
      this._reload().catch(e => this._showError(e));
    }
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onDispose(): void {
    this._themeProvider.themeChangedEvent.remove(this, this._onThemeChanged);
    if (this._root) this._root.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const inline = this.properties.configSource !== 'url';
    const problems = inline && this.properties.configJson
      ? parseConfig(this.properties.configJson).diagnostics.filter(d => d.severity === 'error')
      : [];
    const sourceFields = inline
      ? [
          PropertyPaneButton('openEditor', {
            text: strings.OpenEditorLabel,
            buttonType: PropertyPaneButtonType.Primary,
            icon: 'Code',
            onClick: () => this._setEditorOpen(true),
          }),
          PropertyPaneLabel('problems', {
            text: problems.length ? `${strings.ProblemsLabel}: ${problems.slice(0, 3).map(d => d.message).join('; ')}` : '',
          }),
          PropertyPaneTextField('configJson', {
            label: strings.ConfigJsonLabel,
            multiline: true,
            rows: 14,
            resizable: true,
          }),
        ]
      : [
          PropertyPaneTextField('configUrl', {
            label: strings.ConfigUrlLabel,
            description: strings.ConfigUrlDescription,
            placeholder: '/sites/pmo/SiteAssets/metrics.json',
          }),
          PropertyPaneButton('reload', {
            text: strings.ReloadLabel,
            icon: 'Refresh',
            onClick: () => this._reload().catch(e => this._showError(e)),
          }),
        ];
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ConfigGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('configSource', {
                  label: strings.ConfigSourceLabel,
                  options: [
                    { key: 'inline', text: strings.ConfigSourceInline },
                    { key: 'url', text: strings.ConfigSourceUrl },
                  ],
                }),
                ...sourceFields,
              ],
            },
          ],
        },
      ],
    };
  }

  private _onThemeChanged(args: ThemeChangedEventArgs): void {
    this._theme = args.theme;
    this.render();
  }

  private _setEditorOpen(open: boolean): void {
    this._editorOpen = open;
    this.render();
  }

  private _saveConfig(text: string): void {
    this.properties.configSource = 'inline';
    this.properties.configJson = text;
    this._editorOpen = false;
    this._reload().catch(e => this._showError(e));
  }

  private _showError(e: unknown): void {
    this._configError = e instanceof Error ? e.message : String(e);
    this.render();
  }

  private async _reload(): Promise<void> {
    await this._loadConfig();
    this._reloadToken++;
    this.render();
    if (this.context.propertyPane.isPropertyPaneOpen()) this.context.propertyPane.refresh();
  }

  private async _loadConfig(): Promise<void> {
    this._config = undefined;
    this._configError = undefined;
    if (this.properties.configSource === 'url') {
      const url = (this.properties.configUrl || '').trim();
      if (!url) return;
      const absolute = new URL(url, this.context.pageContext.web.absoluteUrl).toString();
      const response = await this.context.spHttpClient.get(absolute, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json, text/plain' },
      });
      if (!response.ok) {
        this._configError = `${strings.LoadFailed} ${url}: ${await errorMessage(response)}`;
        return;
      }
      this._applyText(await response.text(), url);
      return;
    }
    const text = this.properties.configJson;
    if (text && text.trim()) this._applyText(text, strings.WebPartProperties);
  }

  private _applyText(text: string, origin: string): void {
    const result = parseConfig(text);
    if (result.config) {
      this._config = result.config;
      return;
    }
    const messages = result.diagnostics.filter(d => d.severity === 'error').map(d => `• ${d.message}`);
    this._configError = `${strings.InvalidJson} ${origin}:\n${messages.join('\n')}`;
  }
}

async function errorMessage(response: HttpClientResponse): Promise<string> {
  const body = await response.text();
  let detail = body.slice(0, 200);
  try {
    const json = JSON.parse(body);
    const error = json['odata.error'] || json.error;
    if (error && error.message) detail = typeof error.message === 'string' ? error.message : error.message.value;
  } catch (e) {
    if (!(e instanceof SyntaxError)) throw e;
  }
  return `HTTP ${response.status} ${response.statusText}${detail ? ` – ${detail}` : ''}`;
}
