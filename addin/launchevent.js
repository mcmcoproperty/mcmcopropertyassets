<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <Id>8a4c2d19-7e51-4b83-9c2a-6f1e4d90b3c7</Id>
  <Version>1.1.0.0</Version>
  <ProviderName>McManus &amp; Co Property</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="McManus Signature" />
  <Description DefaultValue="Automatically inserts the McManus &amp; Co Property email signature." />
  <IconUrl DefaultValue="https://mcmcoproperty.github.io/mcmcopropertyassets/brandmark.png" />
  <HighResolutionIconUrl DefaultValue="https://mcmcoproperty.github.io/mcmcopropertyassets/brandmark.png" />
  <SupportUrl DefaultValue="https://mcmanuscoproperty.com.au" />

  <AppDomains>
    <AppDomain>https://mcmcoproperty.github.io</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://mcmcoproperty.github.io/mcmcopropertyassets/addin/autorun.html" />
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>

  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
  </Rule>

  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.5">
          <bt:Set Name="Mailbox" />
        </bt:Sets>
      </Requirements>

      <Hosts>
        <Host xsi:type="MailHost">
          <Runtimes>
            <Runtime resid="WebViewRuntime.Url">
              <Override type="javascript" resid="JSRuntime.Url" />
            </Runtime>
          </Runtimes>

          <DesktopFormFactor>
            <SupportsSharedFolders>true</SupportsSharedFolders>
            <ExtensionPoint xsi:type="LaunchEvent">
              <LaunchEvents>
                <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler" />
                <LaunchEvent Type="OnMessageFromChanged" FunctionName="onMessageFromChangedHandler" />
              </LaunchEvents>
              <SourceLocation resid="WebViewRuntime.Url" />
            </ExtensionPoint>
          </DesktopFormFactor>

          <MobileFormFactor>
            <ExtensionPoint xsi:type="LaunchEvent">
              <LaunchEvents>
                <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler" />
                <LaunchEvent Type="OnMessageFromChanged" FunctionName="onMessageFromChangedHandler" />
              </LaunchEvents>
              <SourceLocation resid="WebViewRuntime.Url" />
            </ExtensionPoint>
          </MobileFormFactor>
        </Host>
      </Hosts>

      <Resources>
        <bt:Urls>
          <bt:Url id="WebViewRuntime.Url" DefaultValue="https://mcmcoproperty.github.io/mcmcopropertyassets/addin/autorun.html" />
          <bt:Url id="JSRuntime.Url" DefaultValue="https://mcmcoproperty.github.io/mcmcopropertyassets/addin/launchevent.js" />
        </bt:Urls>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
