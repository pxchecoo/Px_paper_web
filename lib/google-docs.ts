"use client";

const DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const GOOGLE_DOC_MIME = "application/vnd.google-apps.document";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DEFAULT_GOOGLE_CLIENT_ID = "384206589619-q4osqu9nid7pi8neu1co96g96e58soj5.apps.googleusercontent.com";

type GoogleTokenResponse = {
  access_token?: string;
  error?: string;
  error_description?: string;
};

type GoogleTokenClient = {
  callback: (response: GoogleTokenResponse) => void;
  requestAccessToken: (options?: { prompt?: string }) => void;
};

type GoogleIdentity = {
  accounts: {
    oauth2: {
      initTokenClient: (config: {
        client_id: string;
        scope: string;
        callback: (response: GoogleTokenResponse) => void;
        error_callback?: (error: { type?: string; message?: string }) => void;
      }) => GoogleTokenClient;
    };
  };
};

declare global {
  interface Window {
    google?: GoogleIdentity;
  }
}

let cachedAccessToken: string | null = null;

function googleClientId() {
  return process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID || DEFAULT_GOOGLE_CLIENT_ID;
}

async function waitForGoogleIdentity(timeoutMs = 10_000): Promise<GoogleIdentity> {
  const started = Date.now();
  while (!window.google?.accounts?.oauth2) {
    if (Date.now() - started > timeoutMs) {
      throw new Error("Google sign-in could not load. Check your connection and try again.");
    }
    await new Promise((resolve) => window.setTimeout(resolve, 80));
  }
  return window.google;
}

async function requestDriveToken(forceNew = false): Promise<string> {
  if (cachedAccessToken && !forceNew) return cachedAccessToken;

  const google = await waitForGoogleIdentity();
  const clientId = googleClientId();

  return new Promise<string>((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: DRIVE_FILE_SCOPE,
      callback: (response) => {
        if (response.error || !response.access_token) {
          reject(new Error(response.error_description || response.error || "Google authorization was cancelled."));
          return;
        }
        cachedAccessToken = response.access_token;
        resolve(response.access_token);
      },
      error_callback: (error) => {
        reject(new Error(error.message || error.type || "Google authorization failed."));
      },
    });

    client.requestAccessToken();
  });
}

async function uploadAsGoogleDoc(blob: Blob, fileName: string, accessToken: string) {
  const cleanName = fileName.replace(/\.docx$/i, "") || "PX Paper document";
  const boundary = `px_paper_${crypto.randomUUID().replace(/-/g, "")}`;
  const metadata = {
    name: cleanName,
    mimeType: GOOGLE_DOC_MIME,
  };

  const body = new Blob(
    [
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`,
      `--${boundary}\r\nContent-Type: ${DOCX_MIME}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`,
    ],
    { type: `multipart/related; boundary=${boundary}` },
  );

  const response = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,webViewLink",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
      body,
    },
  );

  if (response.status === 401) {
    cachedAccessToken = null;
    throw new Error("GOOGLE_TOKEN_EXPIRED");
  }

  if (!response.ok) {
    const details = await response.text();
    throw new Error(`Google Drive upload failed (${response.status}). ${details.slice(0, 180)}`);
  }

  const created = (await response.json()) as { id?: string; mimeType?: string };
  if (!created.id) throw new Error("Google created the file but did not return a document ID.");
  return created.id;
}

export async function createGoogleDocFromDocx(blob: Blob, fileName: string) {
  let token = await requestDriveToken();

  try {
    const id = await uploadAsGoogleDoc(blob, fileName, token);
    return `https://docs.google.com/document/d/${encodeURIComponent(id)}/edit`;
  } catch (error) {
    if (error instanceof Error && error.message === "GOOGLE_TOKEN_EXPIRED") {
      token = await requestDriveToken(true);
      const id = await uploadAsGoogleDoc(blob, fileName, token);
      return `https://docs.google.com/document/d/${encodeURIComponent(id)}/edit`;
    }
    throw error;
  }
}
