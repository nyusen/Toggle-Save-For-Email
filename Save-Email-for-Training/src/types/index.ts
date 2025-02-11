export interface EmailData {
  subject: string;
  body: string;
  sender: string;
  recipients: string[];
  timestamp: string;
}

export interface NotificationMessage {
  type: Office.MailboxEnums.ItemNotificationMessageType;
  message: string;
  icon: string;
  persistent: boolean;
}

export interface SaveEmailResponse {
  success: boolean;
  message?: string;
  error?: string;
}
