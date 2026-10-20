import { toast } from "sonner";
import { getInlineAttachmentType } from "./inline-attachment";

export const downloadAttachment = async (attachment: {
  body: string;
  mimeType: string;
  filename: string;
  attachmentId: string;
}) => {
  try {
    const attachmentData = attachment.body;

    const byteCharacters = atob(attachmentData);
    const byteNumbers: number[] = Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: attachment.mimeType });

    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = attachment.filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Error downloading attachment:", error);
    toast.error("Failed to download attachment");
  }
};

export const openAttachmentViewer = () => {
  const width = 800;
  const height = 600;
  const left = (window.screen.width - width) / 2;
  const top = (window.screen.height - height) / 2;
  return window.open(
    "",
    "attachment-viewer",
    `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,location=no,menubar=no`,
  );
};

export const openAttachment = async (
  attachment: {
    body: string;
    mimeType: string;
    filename: string;
    attachmentId: string;
  },
  pendingPopup?: Window | null,
) => {
  const inlineType = getInlineAttachmentType(attachment.mimeType);
  if (!inlineType) {
    pendingPopup?.close();
    await downloadAttachment(attachment);
    return;
  }
  try {
    const attachmentData = attachment.body;

    const byteCharacters = atob(attachmentData);
    const byteNumbers: number[] = Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: inlineType });
    const url = window.URL.createObjectURL(blob);

    const popup = pendingPopup === undefined ? openAttachmentViewer() : pendingPopup;

    if (popup) {
      popup.location.href = url;
      popup.focus();
      // Clean up the URL after a short delay to ensure the browser has time to load it
      setTimeout(() => window.URL.revokeObjectURL(url), 1000);
    } else {
      // Popup was blocked: nothing will ever load the URL, so revoke it now
      // instead of leaking the blob until the tab closes.
      window.URL.revokeObjectURL(url);
    }
  } catch (error) {
    console.error("Error opening attachment:", error);
    toast.error("Failed to open attachment");
  }
};
