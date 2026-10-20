import { openAttachment, openAttachmentViewer } from "../../../email/client/lib/attachment-actions";

Object.assign(window, {
  openMailAttachment: (attachment: Parameters<typeof openAttachment>[0], blocked = false) =>
    openAttachment(attachment, blocked ? null : openAttachmentViewer()),
  pngAttachment: () => {
    const canvas = document.createElement("canvas");
    canvas.width = canvas.height = 1;
    return canvas.toDataURL("image/png").split(",")[1];
  },
});
