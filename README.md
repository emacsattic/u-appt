# u-appt

## Introduction

This package contains commands for extracting appointments from a
buffer, converting them and adding them to an Emacs diary file.  If
an appointment is found, you will be asked, whether you want to add
it to your diary file.  Currently these commands will recognize
appointments that were sent by

- MS Outlo*k (English, German, Norwegian)
- L*tus N*tes (German only)
- a very proprietary Spanish appointment transmitter

In order to use the commands, add the following to your .emacs file:

    (autoload 'u-appt-check-outlook "u-appt" "Check for outlook invitations" t)
    (autoload 'u-appt-check-notes "u-appt" "Check for notes invitations" t)

Say `M-x u-appt-check-outlook` whenever you run across an
appointment which you want to add to your diary.

VM users want to put this into their VM config file:

    (add-hook 'vm-select-new-message-hook 'u-appt-check-outlook)
    (add-hook 'vm-select-new-message-hook 'u-appt-check-notes)

Unfortunately this does not work for mime-encoded messages.  They
are decoded **after** that hook is run.  So you have to call these
functions by hand, or you could use

    (defadvice vm-decode-mime-message (after u-appt activate)
           (u-appt-check-outlook))
    (defadvice vm-decode-mime-message (after u-appt activate)
           (u-appt-check-notes))

but that would call the functions each time you look at a
message.  Bad idea.

Gnus users might want to use the following.

    (defun my-gnus-check-outlook ()
      "Run from a hook to check new messages in Gnus for Outlook appointment
    invitations, and offer to save them in the diary."
      (save-excursion
        (let ((mark (gnus-summary-article-mark)))
          (when (gnus-unread-mark-p mark)
         (set-buffer gnus-article-buffer)
         (u-appt-check-outlook)))))
    (add-hook 'gnus-mark-article-hook 'my-gnus-check-outlook)
    ;; Add the original value of gnus-mark-article-hook, since this
    ;; is overwritten by the above add-hook.
    ;; Add it at the end, since it sets the mark to read.
    (add-hook 'gnus-mark-article-hook
    'gnus-summary-mark-read-and-unread-as-read t)


## History

* **0.5** Swedish Outlook improved, by Cristian Ionescu-Idbohrn. (2005-01-22)

* **0.4** Swedish Outlook, from Cristian Ionescu-Idbohrn. (2004-10-30)

* **0.3** Norwegian Outlook (i.e. US Outlook with Norwegian locale), from
Steinar Bang.  (2004-10-09)

* **0.2** Bugfixes, thanks to Colin Marquardt.

* **0.1** First version.
