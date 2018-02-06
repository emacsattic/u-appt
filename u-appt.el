;;; u-appt.el --- Appointment parser                    -*- coding: utf-8 -*-

;;  Copyright (C) 2002, 2004 by Ulf Jasper

;;  Author:     Ulf Jasper <ulf.jasper@web.de>
;;  Keywords:   diary calendar outlook lotus
;;  Time-stamp: "22. Januar 2005, 19:44:14 (ulf)"
;;  Version:    0.5
;;  CVS:        $Id$

;; ======================================================================

;;  This program is free software; you can redistribute it and/or modify it
;;  under the terms of the GNU General Public License as published by the
;;  Free Software Foundation; either version 2 of the License, or (at your
;;  option) any later version.

;;  This program is distributed in the hope that it will be useful, but
;;  WITHOUT ANY WARRANTY; without even the implied warranty of
;;  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
;;  General Public License for more details.

;;  You should have received a copy of the GNU General Public License along
;;  with this program; if not, write to the Free Software Foundation, Inc.,
;;  59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

;; ======================================================================
;;; Commentary:

;;  This package contains commands for extracting appointments from a
;;  buffer, converting them and adding them to an Emacs diary file.  If
;;  an appointment is found, you will be asked, whether you want to add
;;  it to your diary file.  Currently these commands will recognize
;;  appointments that were sent by

;;  - MS Outlo*k (English, German, Norwegian)
;;  - L*tus N*tes (German only)
;;  - a very proprietary Spanish appointment transmitter

;;  In order to use the commands, add the following to your .emacs file:

;;  (autoload 'u-appt-check-outlook "u-appt" "Check for outlook invitations" t)
;;  (autoload 'u-appt-check-notes "u-appt" "Check for notes invitations" t)

;;  Say `M-x u-appt-check-outlook' whenever you run across an
;;  appointment which you want to add to your diary.

;;  VM users want to put this into their VM config file:

;;   (add-hook 'vm-select-new-message-hook 'u-appt-check-outlook)
;;   (add-hook 'vm-select-new-message-hook 'u-appt-check-notes)

;;  Unfortunately this does not work for mime-encoded messages.  They
;;  are decoded *after* that hook is run.  So you have to call these
;;  functions by hand, or you could use

;;  (defadvice vm-decode-mime-message (after u-appt activate)
;;         (u-appt-check-outlook))
;;  (defadvice vm-decode-mime-message (after u-appt activate)
;;         (u-appt-check-notes))

;;  but that would call the functions each time you look at a
;;  message.  Bad idea.

;;  Gnus users might want to use the following.

;;  (defun my-gnus-check-outlook ()
;;    "Run from a hook to check new messages in Gnus for Outlook appointment
;;  invitations, and offer to save them in the diary."
;;    (save-excursion
;;      (let ((mark (gnus-summary-article-mark)))
;;        (when (gnus-unread-mark-p mark)
;;       (set-buffer gnus-article-buffer)
;;       (u-appt-check-outlook)))))
;;  (add-hook 'gnus-mark-article-hook 'my-gnus-check-outlook)
;;  ;; Add the original value of gnus-mark-article-hook, since this
;;  ;; is overwritten by the above add-hook.
;;  ;; Add it at the end, since it sets the mark to read.
;;  (add-hook 'gnus-mark-article-hook
;;  'gnus-summary-mark-read-and-unread-as-read t)

;; ======================================================================
;;; History:

;;  0.5: Swedish Outlook improved, by Cristian Ionescu-Idbohrn.
;;       (2005-01-22)

;;  0.4: Swedish Outlook, from Cristian Ionescu-Idbohrn. (2004-10-30)

;;  0.3: Norwegian Outlook (i.e. US Outlook with Norwegian locale), from
;;       Steinar Bang.  (2004-10-09)

;;  0.2: Bugfixes, thanks to Colin Marquardt.

;;  0.1: First version.

;; ======================================================================
;;; Code:

(require 'appt)
(require 'mailheader)

(defconst u-appt-monthnumber-table
  '(("^jan\\(uar\\)?[iy]?$"          . 1)
    ("^feb\\(ruar\\)?[iy]?$"         . 2)
    ("^mar\\(ch\\|s\\)?\\|märz?$" . 3)
    ("^apr\\(il\\)?$"             . 4)
    ("^ma[ijy]$"                   . 5)
    ("^jun[ie]?$"                 . 6)
    ("^jul[iy]?$"                 . 7)
    ("^aug\\(ust\\)?i?$"            . 8)
    ("^sep\\(tember\\)?$"         . 9)
    ("^o[ck]t\\(ober\\)?$"        . 10)
    ("^nov\\(ember\\)?$"          . 11)
    ("^de[czs]\\(ember\\)?$"      . 12))
  "Regexps for month names.  Currently only German, English,
Norwegian, and Swedish.")

(defun u-appt-handle (subject string)
  "Asks user whether to add an appointment.
SUBJECT is the appointment-subject
STRING is the formatted diary entry"
  (if (y-or-n-p (format "Add appointment for `%s' to diary? " subject))
      (save-window-excursion
        (make-diary-entry string)
        (save-excursion
          (set-buffer (find-buffer-visiting diary-file))
          (save-buffer))
        ;; hmmm... FIXME!
        (if (fboundp 'appt-initialize)
            (appt-initialize))
        (if (fboundp 'appt-activate)
            (appt-activate 1))
        (message "Addded %s to diary" string))))

(defun u-appt-date-string (date &optional abbreviate nodayname)
  "Return properly formatted DATE.
Takes care of optional arguments ABBREVIATE and NODAYNAME."
  (let ((calendar-date-display-form (if european-calendar-style
                                        '(day " " monthname " " year)
                                      '(monthname " " day " " year))))
    (calendar-date-string date abbreviate nodayname)))

(defsubst u-appt-get-month-number (monthname)
  "Return the month number for the given MONTHNAME."
  (save-match-data
    (let ((case-fold-search t))
      (assoc-default monthname u-appt-monthnumber-table 'string-match))))


(defun u-appt-check-outlook (&rest args)
  "Search a buffer for an Outlook-style appointment and add a diary entry.
Optional argument ARGS is unused!"
  (interactive)
  (let (subject day month year time am-pm string type header-list
		(appt-found nil)
		(where ""))
    (save-excursion
      (goto-char (point-min))
      ;; parse mail headers
      (setq header-list (mail-header-extract))
      ;; find the subject
      (if header-list
	  (progn
	    (setq subject (mail-header 'subject header-list))
	    (setq type    (mail-header 'content-type header-list)))
	;; fall-back, in case that mail-header-extract failed
	(progn
	  (goto-char (point-min))
	  (if (re-search-forward "^Subject:\\s-+\\(.*\\)$" nil t)
	      (setq subject (match-string-no-properties 1)))))
      ;; find the location
      (if (re-search-forward "^Where:\\s-+\\(.*\\)$" nil t)
	  (setq where (match-string-no-properties 1)))
      (goto-char (point-min))
      ;; 
      (if (or (not type) (not (string-match "message" type)))
          (progn
            (cond (;; German
                   ;; Example:
                   ;; Zeit: Freitag, 6. Dezember 2002 00:07 Stiefel rausstellen!
                   ;; Also seen:
                   ;; When: Freitag, 4. Juni 2004 00:00 Termin!
                   (re-search-forward
                    (concat "^\\(Zeit\\|When\\): [^ ]+, +\\([0-9]+\\)\. +"
                            "\\([A-Z][a-zäöü][äöüa-z]\\)[^ ]* +\\([0-9]+\\) +"
                            "\\([^ ]+\\)\\s-+\\(.*\\)$") nil t)
                   (setq day   (match-string-no-properties 2)
			 month (match-string-no-properties 3)
			 year  (match-string-no-properties 4)
			 time  (match-string-no-properties 5))
                   ;; This is probably NOT the subject:
                   ;;(if (> (length (match-string-no-properties 6)) 0)
                   ;;  (setq subject (match-string-no-properties 6)))
                   (setq string (format "%s %s %s"
                                        (u-appt-date-string
                                         (list (u-appt-get-month-number month)
                                               (string-to-number day)
                                               (string-to-number year)) t t)
                                        time subject))
                   (setq appt-found t))
                  (;; English
                   ;; Example: FIXME
                   (re-search-forward
                    (concat "^\\(Start Date\\):\\s-+\\([0-9]+\\)/\\([0-9]+\\)/"
                            "\\([0-9]+\\)\\s-+"
                            "\\([0-9]+:[0-9]+\\)\\s-*\\([ap]m\\)"
                            "\\(.*\\)$") nil t)
                   (setq
		    month (string-to-number (match-string-no-properties 2))
		    day   (string-to-number (match-string-no-properties 3))
		    year  (string-to-number (match-string-no-properties 4))
		    time  (match-string-no-properties 5)
		    am-pm (match-string-no-properties 6))
                   (if (> (length (match-string-no-properties 7)) 0)
                       (setq subject (match-string-no-properties 7)))
                   (setq string (format "%s %s %s %s"
                                        (u-appt-date-string
                                         (list month day year) t t)
                                        time am-pm subject))
                   (setq appt-found t))
                  (;; US Outlook2003 with Norwegian date setting
                   ;; Example (note linebreak):
                   ;; When: 2. september 2004 10:00-11:30 (GMT+01:00) Amsterdam, Berlin, Bern,
                   ;; Rome, Stockholm, Vienna.
                   (re-search-forward
                    (concat "^\\(When\\):\\s-+\\([0-9]+\\)\\.\\s-+"
                            "\\([A-Za-z]+\\)\\s-+"
                            "\\([0-9]+\\)\\s-+"
                            "\\([0-9]+:[0-9]+\\)"
                            "\\(.*\\)$") nil t)
                   (setq day   (match-string-no-properties 2)
			 month (match-string-no-properties 3)
			 year  (match-string-no-properties 4)
			 time  (match-string-no-properties 5))
                   (setq string (format "%s %s %s"
                                        (u-appt-date-string
                                         (list (u-appt-get-month-number month)
                                               (string-to-number day)
                                               (string-to-number year)) t t)
                                        time subject))
                   (setq appt-found t))
                  (;; US Outlook2003 with US long date and no am/pm
                   ;; US long date is dddd, MMMM dd, yyyy
                   ;; Example (note linebreak):
                   ;; When: Monday, October 11, 2004 16:00-16:05 (GMT+01:00) Amsterdam, Berlin, Bern,
                   ;; Rome, Stockholm, Vienna.
                   (re-search-forward
                    (concat "^\\(When\\):\\s-+\\([A-Za-z]+\\),\\s-+"
                            "\\([A-Za-z]+\\)\\s-+"
                            "\\([0-9]+\\),\\s-+"
                            "\\([0-9]+\\)\\s-+"
                            "\\([0-9]+:[0-9]+\\)"
                            "\\(.*\\)$") nil t)
                   (setq month (match-string-no-properties 3)
			 day   (match-string-no-properties 4)
			 year  (match-string-no-properties 5)
			 time  (match-string-no-properties 6))
                   (setq string (format "%s %s %s"
                                        (u-appt-date-string
                                         (list (u-appt-get-month-number month)
                                               (string-to-number day)
                                               (string-to-number year)) t t)
                                        time subject))
                   (setq appt-found t))
		  (;; US Outlook with SE long date and no am/pm
		   ;; Example (note linebreak):
		   ;; When: den 22 oktober 2004 13:30-14:30 (GMT+01:00) Amsterdam, Berlin,
		   ;; Bern, Rome, Stockholm, Vienna.
		   (re-search-forward
		    (concat "^\\(When\\):\\s-+\\([A-Za-z]+\\)\\s-+" ;; match 1,2
			    "\\([0-9]+\\)\\s-+"	   ;; day, match 3
			    "\\([A-Za-z]+\\)\\s-+" ;; month, match 4
			    "\\([0-9]+\\)\\s-+"	   ;; year, match 5
			    "\\([0-9]+:[0-9]+\\)"  ;; time, match 6
			    "\\(.*\\)$") nil t)
		   (setq day	  (match-string-no-properties 3)
			 month (match-string-no-properties 4)
			 year  (match-string-no-properties 5)
			 time  (match-string-no-properties 6))
		   (setq string (format "%s\n %s %s %s\n"
					(u-appt-date-string
					 (list (u-appt-get-month-number month)
					       (string-to-number day)
					       (string-to-number year)) t t)
					time
					subject
					where))
		   (setq appt-found t))
                  (;; Spanish -- not really outlook but similar
                   ;; Example:
                   ;; Fecha: miércoles 16 de octubre 2002 10:00am  convocatoria
                   (re-search-forward
                    (concat "^\\(Fecha\\): [^ ]+ +\\([0-9]+\\) +[^ ]+ +"
                            "\\([A-Za-z][a-z][a-z]\\)[^ ]* +\\([0-9]+\\) +"
                            "\\([0-9]+:[0-9]+\\)\\s-*\\([ap]m\\)\\s-+"
                            "\\(.*\\)$") nil t)
                   (setq day   (match-string-no-properties 2)
			 month (capitalize (match-string-no-properties 3))
			 year  (match-string-no-properties 4)
			 time  (match-string-no-properties 5)
			 am-pm (match-string-no-properties 6))
                   (if (> (length (match-string-no-properties 7)) 0)
                       (setq subject (match-string-no-properties 7)))
                   ;; CHECKME: does this work?
                   (setq string (format "%s %s %s"
                                        (u-appt-date-string
                                         (list (u-appt-get-month-number month)
                                               (string-to-number day)
                                               (string-to-number year)) t t)
                                        time subject))
                   ;;(setq string (format "%s %s %s %s%s %s"
                   ;;day month year time am-pm subject))
                   (setq appt-found t))))))
    (when appt-found
      (u-appt-handle subject string))))

(defun u-appt-check-notes (&rest args)
  "Search a buffer for a Lotus-Notes-style appointment and add to diary.
Optional argument ARGS is unused!"
  (interactive)
  (let (subject start-day start-month start-year start-time end-day
                end-month end-year end-time string type header-list
                (appt-found nil)
                (has-end nil))
    (save-excursion
      (goto-char (point-min))
      (setq header-list (mail-header-extract))
      (when header-list
        (setq subject (mail-header 'subject header-list))
        (setq type    (mail-header 'content-type header-list)))
      (when (or (not type) (not (string-match "message" type)))
        (when (re-search-forward "^\\s-*Calendar Entry:\\s-*" nil t)
          (when (re-search-forward
                 (concat "^\\s-*Begins:\\s-+\\([0-9]+\\).\\([0-9]+\\)."
                         "\\([0-9]+\\)\\s-+\\([0-9]+:[0-9]+\\)\\s-*"
                         ".*$") nil t)
            (setq
             start-day   (string-to-number (match-string-no-properties 1))
             start-month (string-to-number (match-string-no-properties 2))
             start-year  (string-to-number (match-string-no-properties 3))
             start-time  (match-string-no-properties 4))
            (setq string (format "%s %s %s"
                                 (u-appt-date-string
                                  (list start-month start-day start-year)
                                  t t)
                                 start-time subject))
            (setq appt-found t))
          (when (re-search-forward
                 (concat "^\\s-*Ends:\\s-+\\([0-9]+\\).\\([0-9]+\\)."
                         "\\([0-9]+\\)\\s-+\\([0-9]+:[0-9]+\\)\\s-*"
                         ".*$") nil t)
            (setq
             end-day   (string-to-number (match-string-no-properties 1))
             end-month (string-to-number (match-string-no-properties 2))
             end-year  (string-to-number (match-string-no-properties 3))
             end-time  (match-string-no-properties 4))
            (if (and (eq start-day end-day)
                     (eq start-month end-month)
                     (eq start-year end-year))
                (setq string (format "%s %s-%s %s"
                                     (u-appt-date-string
                                      (list start-month start-day start-year)
                                      t t)
                                     start-time end-time subject))
              (setq string (format "%%%%(diary-block %s %s) %s"
                                   (u-appt-date-string
                                    (list start-month start-day start-year)
                                    t t)
                                   (u-appt-date-string
                                    (list end-month end-day end-year)
                                    t t)
                                   subject)))
            (setq has-end t)))))
    (when appt-found
      (u-appt-handle subject string))))

(provide 'u-appt)
;;; u-appt.el ends here
