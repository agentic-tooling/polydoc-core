/**
 * Imports all three published entry points in one process and reports whether
 * the transport contract they share is one binding or several copies.
 *
 * The root barrel and both transport entry points re-export `TransportError`
 * and `DOCX_MIME_TYPE`. That is only safe if every entry point resolves to the
 * same `dist/transport.js` module instance — otherwise `instanceof` would fail
 * across entry points and the duplicate exports would be a real hazard.
 */
import * as core from "@llbbl/polydoc-core";
import * as google from "@llbbl/polydoc-core/google";
import * as sharepoint from "@llbbl/polydoc-core/sharepoint";

const googleError = new google.GoogleDriveTransportError(
  "GOOGLE_DRIVE_API_FAILED",
  "probe",
  "probe",
);
const sharePointError = new sharepoint.SharePointTransportError(
  "SHAREPOINT_HTTP_FAILED",
  "probe",
  "probe",
);

process.stdout.write(
  JSON.stringify({
    googleReExportsSameTransportError: google.TransportError === core.TransportError,
    sharePointReExportsSameTransportError: sharepoint.TransportError === core.TransportError,
    googleReExportsSameDocxMimeType: google.DOCX_MIME_TYPE === core.DOCX_MIME_TYPE,
    sharePointReExportsSameDocxMimeType: sharepoint.DOCX_MIME_TYPE === core.DOCX_MIME_TYPE,
    googleErrorIsCoreTransportError: googleError instanceof core.TransportError,
    sharePointErrorIsCoreTransportError: sharePointError instanceof core.TransportError,
  }),
);
