var testEmail = {
    subject: 'Test Subject',
    body: 'Test Body',
    prependBody: 'Pre-',
    setSelected: '-setSelected',
    fileAttachment: {
        name: 'File Attachment',
        file: 'http://i.imgur.com/ihH0kfg.jpg'
    },
    itemAttachmentId: null,
    recipients: {
        to: ['a@example.com', 'b@example.com'],
        cc: ['c@example.com', 'd@example.com'],
        bcc: ['e@example.com', 'f@example.com'],
        extra: ['deleteme@example.com']

    },
    appointment: {
        startTime: new Date(2015, 11, 18, 0, 0, 0),
        endTime: new Date(2015, 11, 18, 2, 30, 0),
        location: 'Cinerama',
        required: ['a@example.com', 'b@example.com'],
        optional: ['c@example.com', 'd@example.com'],
        extra: ['deleteme@example.com']
    },
    messageKey: "msg",
    firstMessage: {
        type: "progressIndicator",
        message: "thinking"

    },
    secondMessage: {
        type: "errorMessage",
        message: "forgot"
    }
};
