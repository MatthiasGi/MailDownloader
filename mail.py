from imapclient import IMAPClient
import email
from datetime import datetime
import os
import email.generator
import re

class Mail:
    """
    Wrapper for all needed methods to process the mails on the IMAP-server.

    Attributes
    ----------
    config : config.Config
        Configuration manager wrapping around the yaml-config-file.
    """

    # List of all Content-types of attachments that should be saved seperatly
    valid_ctypes = ['application/pdf']

    # Regexp that matches all allowed characters in a filename
    filesafe = re.compile(r'[\w\-\. ]')

    def __init__(self, config):
        config.checkParams('server', 'port', 'username', 'password', 'inbox',
                           'outbox', 'basepath', 'eml-to-pdf-path')
        host = config.get('server')
        port = config.get('port')
        username = config.get('username')
        password = config.get('password')
        inbox = config.get('inbox')
        self.outbox = config.get('outbox')
        self.basepath = config.get('basepath')
        if not os.path.exists(self.basepath): os.makedirs(self.basepath)
        self.emltopdf = config.get('eml-to-pdf-path')

        self.server = IMAPClient(host, port=port)
        result = self.server.login(username, password)
        print(result.decode('UTF-8'))
        self.server.select_folder(inbox)

    def check(self):
        """
        Checks for new mail to process on the server.
        """
        messages = self.server.search(['ALL'])
        for id, data in self.server.fetch(messages, 'RFC822').items():
            mail = email.message_from_bytes(data[b'RFC822'])
            self.processMail(mail)

        self.server.copy(messages, self.outbox)
        self.server.delete_messages(messages)
        self.server.expunge()

    def getDate(self, mail):
        """
        Extracts the date from the mail header and rewrites it for use as the
        beginning of a filename.

        Parameters
        ----------
        mail : email
            Mail from which to extract the date.

        Returns
        -------
        The date reformatted to be used in a filename.
        """
        date = mail.get('Date')
        date = datetime.strptime(date[:-6], '%a, %d %b %Y %H:%M:%S')
        return date.strftime('%Y%m%d-%H%M%S')

    def processMail(self, mail):
        """
        Processes a specific mail by saving it.

        Parameters
        ----------
        mail : email
            The mail to process.
        """
        filename = '%s-%s.eml' % (self.getDate(mail), mail.get('Subject'))
        filename = self.validateFilename(filename)
        filename = os.path.join(self.basepath, filename)
        with open(filename, 'w') as f:
            generator = email.generator.Generator(f)
            generator.flatten(mail)

        if mail.is_multipart(): self.processAttachments(mail)

    def processAttachments(self, mail):
        """
        Processes all attachments of a mail.

        Parameters
        ----------
        mail : email
            The mail to process.
        """
        date = self.getDate(mail)
        for part in mail.walk():
            ctype = part.get_content_type()
            if not ctype in Mail.valid_ctypes: continue
            filename = '%s-attachment-%s' % (date, part.get_filename())
            filename = self.validateFilename(filename)
            filename = os.path.join(self.basepath, filename)
            with open(filename, 'wb') as f:
                f.write(part.get_payload(decode=True))

    def validateFilename(self, filename):
        """
        Makes a filename safe.

        Parameters
        ----------
        filename : str
            Filename to reformat.

        Returns
        -------
        The new filename, stripped of unsafe characters.
        """
        return ''.join([c for c in filename if Mail.filesafe.match(c)])

    def __del__(self):
        """
        Deconstructor disconnects from the IMAP-Server.
        """
        self.server.logout()
