import { Fragment } from "react";

import { SIGN_OFF_UNIT, type Unit } from "@/lib/application-data";
import type { DateFormat, PaperSize, Tone } from "@/generated/prisma/enums";
import { prepareLetter } from "@/lib/document-prep";
import type { Basics } from "@/lib/profile";
import type { ThemeData } from "@/lib/themes";

import { ContactLine, DocumentPage, plainUnit, type UnitRenderer } from "./document-page";
import styles from "./document.module.css";

type CoverLetterDocumentProps = {
  theme: ThemeData;
  paper: PaperSize;
  /** The sender, for a letter without its own sender and signature units. */
  basics: Basics;
  /** The letter's parts: frame units, paragraphs ("paragraph:1", …) and its "sign-off" (E8-11). */
  units: Unit[];
  company: string;
  location: string;
  tone: Tone;
  /** The date for a letter without its own date unit, at UTC midnight. */
  date: Date;
  dateFormat: DateFormat;
  /** Overrides the "sign-off" unit; without either, the tone's sign-off. */
  signOff?: string;
  /** The language it's written in (SET-7); without one, it's read from the text. */
  language?: string | null;
  /** Draws each part; the preview's adds editing (E8-3). */
  renderUnit?: UnitRenderer;
};

/**
 * A cover letter (APP-10) in the theme's fonts, colours and spacing, always
 * in a single column: sender, date, recipient, greeting, paragraphs, sign-off
 * and the name under it. Every part is a unit (E8-11); a letter written
 * before that shows the parts it lacks as letters always had them.
 */
export function CoverLetterDocument({
  theme,
  paper,
  basics,
  units,
  company,
  location,
  tone,
  date,
  dateFormat,
  signOff,
  language,
  renderUnit = plainUnit,
}: CoverLetterDocumentProps) {
  const letter = prepareLetter({ theme, paper, basics, units, company, location, tone, date, dateFormat, signOff, language });

  return (
    <DocumentPage theme={letter.resolved} paper={paper} lang={letter.lang} dir={letter.dir} label="Cover letter" className={styles.letter}>
      {renderUnit({
        id: "sender",
        className: styles.sender,
        children: (
          <>
            <p className={styles.senderName}>{letter.senderName}</p>
            {letter.contactLines.map((items, index) => (
              <ContactLine key={index} items={items} />
            ))}
          </>
        ),
      })}
      {renderUnit({ id: "date", className: styles.letterDate, children: letter.date })}
      {renderUnit({
        id: "recipient",
        className: styles.recipient,
        children: letter.recipient.map((line, index) => (
          <p key={index}>{line}</p>
        )),
      })}
      {renderUnit({ id: "greeting", className: styles.greeting, children: letter.greeting })}
      <div className={styles.paragraphs}>
        {letter.paragraphs.map((unit) => (
          <Fragment key={unit.id}>{renderUnit({ id: unit.id, className: styles.paragraph, children: unit.text })}</Fragment>
        ))}
      </div>
      <div className={styles.signOff}>
        {renderUnit({ id: SIGN_OFF_UNIT, children: letter.closing })}
        {renderUnit({ id: "signature", children: letter.signature })}
      </div>
    </DocumentPage>
  );
}
