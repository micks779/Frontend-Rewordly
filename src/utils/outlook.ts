export const copyToClipboard = async (text: string): Promise<void> => {
  try {
    await navigator.clipboard.writeText(text);
  } catch (err) {
    throw new Error('Failed to copy to clipboard');
  }
};

export const replaceInOutlook = async (text: string): Promise<void> => {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      throw new Error('No email selected');
    }

    // Replace the entire body with the text
    item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        throw new Error('Failed to replace text in Outlook');
      }
    });
  } catch (err) {
    throw new Error(err instanceof Error ? err.message : 'Failed to replace text in Outlook');
  }
}; 