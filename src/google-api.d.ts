// Minimal type declarations for Google APIs used in this app

declare namespace gapi {
  function load(apiName: string, callback: () => void): void;

  namespace client {
    function init(config: { discoveryDocs?: string[] }): Promise<void>;
    function setToken(token: { access_token: string }): void;
    function getToken(): { access_token: string } | null;

    function request(config: {
      path: string;
      method: string;
      params?: Record<string, string>;
      headers?: Record<string, string>;
      body?: string;
    }): Promise<{ result: Record<string, unknown>; body: string }>;

    namespace drive {
      namespace files {
        function list(params: {
          q?: string;
          fields?: string;
          pageSize?: number;
          orderBy?: string;
        }): Promise<{
          result: {
            files?: Array<{ id: string; name: string; mimeType?: string }>;
          };
        }>;

        function get(params: {
          fileId: string;
          alt?: string;
        }): Promise<{ result: Record<string, unknown>; body: string }>;

        function create(params: {
          resource: Record<string, unknown>;
          fields?: string;
        }): Promise<{ result: { id: string } }>;

        function update(params: {
          fileId: string;
          resource?: Record<string, unknown>;
          fields?: string;
        }): Promise<{ result: Record<string, unknown> }>;
      }
    }
  }
}

declare namespace google {
  namespace picker {
    class PickerBuilder {
      addView(viewOrId: View | ViewId): PickerBuilder;
      setOAuthToken(token: string): PickerBuilder;
      setDeveloperKey(key: string): PickerBuilder;
      setAppId(appId: string): PickerBuilder;
      setCallback(callback: (data: ResponseObject) => void): PickerBuilder;
      enableFeature(feature: Feature): PickerBuilder;
      disableFeature(feature: Feature): PickerBuilder;
      setTitle(title: string): PickerBuilder;
      setSize(width: number, height: number): PickerBuilder;
      build(): Picker;
    }

    class Picker {
      setVisible(visible: boolean): void;
      dispose(): void;
    }

    class View {
      constructor(viewId: ViewId);
      setMimeTypes(mimeTypes: string): void;
    }

    class DocsView {
      constructor(viewId?: ViewId);
      setIncludeFolders(include: boolean): DocsView;
      setSelectFolderEnabled(enabled: boolean): DocsView;
      setMimeTypes(mimeTypes: string): DocsView;
      setMode(mode: DocsViewMode): DocsView;
    }

    enum ViewId {
      DOCS = 'all',
      DOCS_IMAGES = 'docs-images',
      DOCS_VIDEOS = 'docs-videos',
      DOCUMENTS = 'documents',
      DRAWINGS = 'drawings',
      FOLDERS = 'folders',
      FORMS = 'forms',
      PDFS = 'pdfs',
      PRESENTATIONS = 'presentations',
      SPREADSHEETS = 'spreadsheets',
    }

    enum DocsViewMode {
      GRID = 'grid',
      LIST = 'list',
    }

    enum Feature {
      MINE_ONLY = 'mine-only',
      MULTISELECT_ENABLED = 'multiselectEnabled',
      NAV_HIDDEN = 'navHidden',
      SUPPORT_DRIVES = 'supportDrives',
    }

    enum Action {
      CANCEL = 'cancel',
      PICKED = 'picked',
    }

    enum Response {
      ACTION = 'action',
      DOCUMENTS = 'docs',
      PARENTS = 'parents',
      VIEW = 'viewToken',
    }

    enum Document {
      ID = 'id',
      NAME = 'name',
      URL = 'url',
      MIME_TYPE = 'mimeType',
    }

    interface ResponseObject {
      [Response.ACTION]: Action;
      [Response.DOCUMENTS]?: Array<{
        [Document.ID]: string;
        [Document.NAME]: string;
        [Document.URL]: string;
        [Document.MIME_TYPE]: string;
      }>;
    }
  }

  namespace accounts {
    namespace oauth2 {
      interface TokenResponse {
        access_token: string;
        error?: string;
        expires_in: number;
        scope: string;
        token_type: string;
      }

      interface TokenClient {
        requestAccessToken(config?: { prompt?: string }): void;
      }

      function initTokenClient(config: {
        client_id: string;
        scope: string;
        callback: (response: TokenResponse) => void;
      }): TokenClient;

      function revoke(token: string, callback: () => void): void;
    }
  }
}
