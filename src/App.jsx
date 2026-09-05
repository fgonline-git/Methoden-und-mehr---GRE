import { useState, useRef, useEffect } from "react";
import Papa from "papaparse";
import mammoth from "mammoth";
import { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel } from "docx";
import { useEditor, EditorContent, ReactNodeViewRenderer, NodeViewWrapper } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import { Image as TiptapImage } from "@tiptap/extension-image";
import { TextStyle } from "@tiptap/extension-text-style";
import { Color } from "@tiptap/extension-color";
import Placeholder from "@tiptap/extension-placeholder";
import * as db from "./datenbank";
// Word-Import/-Export für die Methodenbeschreibung ist (noch) nicht enthalten – die dafür
// nötigen Zusatzbibliotheken laufen in dieser Vorschau-Umgebung nicht. Bei Bedarf im eigenen
// Projekt ergänzbar (siehe frühere Chat-Nachrichten für den genauen Code).

// ---------- Design tokens ----------
const T = {
  ink: "#20242C",
  paper: "#F5F4EF",
  paperAlt: "#ECEAE2",
  board: "#1B3A3F",
  boardLight: "#28494E",
  boardActive: "#2C5C60",
  boardText: "#E9EEEA",
  boardTextMuted: "#9FB4AE",
  accent: "#C98A2B",
  accentSoft: "#F0E1BE",
  danger: "#B14B3E",
  dangerSoft: "#F1DAD5",
  success: "#3E7A63",
  successSoft: "#DAE9E1",
  grey: "#8F8B7F",
  greySoft: "#E7E3D7",
  line: "#DBD7CA",
  muted: "#6B6F76",
};

const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
.mc-display{font-family:'Space Grotesk',sans-serif;}
.mc-body{font-family:'Inter',sans-serif;}
.mc-mono{font-family:'IBM Plex Mono',monospace;}
.rt-content:focus{outline:2px solid #C98A2B33;}
.rt-content ul{list-style:disc;margin-left:1.1em;}
.rt-content ol{list-style:decimal;margin-left:1.1em;}
.rt-content strong{font-weight:600;}
.rt-content h2{font-family:'Space Grotesk',sans-serif;font-size:1.15em;font-weight:600;margin:0.5em 0 0.3em;}
.rt-content h3{font-family:'Space Grotesk',sans-serif;font-size:1.05em;font-weight:600;margin:0.4em 0 0.25em;}
.rt-content img{max-width:100%;border-radius:6px;margin:0.4em 0;}
.rt-content h1,.rt-content h4{font-family:'Space Grotesk',sans-serif;font-weight:600;margin:0.4em 0 0.25em;}
.rt-content p{margin:0 0 0.4em;}
.rt-content p.is-editor-empty:first-child:before{content:attr(data-placeholder);color:#B4AF9F;float:left;height:0;pointer-events:none;}
`;

const JAHRGAENGE = [5, 6, 7, 8, 9, 10];
const BUCHSTABEN = ["a", "b", "c"];
const QUARTALE = [1, 2, 3, 4];
const QUARTAL_SPANNE = { 1: "Aug – Okt", 2: "Nov – Jan", 3: "Feb – Apr", 4: "Mai – Jul" };
const STATUS_LABEL = { ausstehend: "ausstehend", erledigt: "erledigt", ausgefallen: "ausgefallen" };

let uidCounter = 1000;
const uid = (p) => `${p}${uidCounter++}`;

// ---------- Seed data ----------
function seedFaecher() {
  // Bewusst leer: Fächer werden real über die Verwaltung angelegt oder aus Untis importiert.
  return [];
}

function seedLehrer() {
  // Bewusst leer: Lehrkräfte werden real über die Verwaltung angelegt oder aus Untis importiert.
  return [];
}

function seedKlassen(lehrer) {
  const klassen = [];
  let i = 0;
  JAHRGAENGE.forEach((jg) => {
    BUCHSTABEN.forEach((b) => {
      klassen.push({
        id: uid("k"),
        jahrgang: jg,
        buchstabe: b,
        lehrer1: lehrer.length ? lehrer[i % lehrer.length].id : null,
        lehrer2: lehrer.length ? lehrer[(i + 1) % lehrer.length].id : null,
        vorgaenger: null,
      });
      i++;
    });
  });
  return klassen;
}

function seedLerngruppen() {
  // Bewusst leer: Lerngruppen setzen reale Fächer/Lehrkräfte voraus (siehe seedFaecher/seedLehrer)
  // und werden entweder händisch in der Verwaltung oder per Untis-Import angelegt.
  return [];
}

// ---------------------------------------------------------------------------
// TEMPORÄR: eingebetteter Untis-Export, wird beim Start automatisch importiert,
// damit die Datei nicht bei jedem Neuladen erneut hochgeladen werden muss.
// Vor dem produktiven Einsatz entfernen (siehe seedInitialerZustand unten)!
// ---------------------------------------------------------------------------
const STANDARD_UNTIS_EXPORT = `lessonId	lessonNumber	subject	teacher	klassen	studentgroup	periods	startDate	endDate	room	foreignKey
32335	44900	BER	WEB			1	02.09.2026	18.07.2027		
32338	45000	BER	WES			1	02.09.2026	18.07.2027		
32347	45300	BER	WÖR			1	02.09.2026	18.07.2027		
33901	40800	BER	BOR			1	02.09.2026	18.07.2027		
33910	41200	BER	EMT			1	02.09.2026	18.07.2027		
33919	41600	BER	GUS			1	02.09.2026	18.07.2027		
33922	41700	BER	HIM			1	02.09.2026	18.07.2027		
33925	41800	BER	HIN			1	02.09.2026	18.07.2027		
33928	41900	BER	HOM			1	02.09.2026	18.07.2027		
33931	42100	BER	KAT			1	02.09.2026	18.07.2027		
33937	42300	BER	FRY			1	02.09.2026	18.07.2027		
33940	42400	BER	KLN			1	02.09.2026	18.07.2027		
33943	42600	BER	KLU			1	02.09.2026	18.07.2027		
33946	42700	BER	KRA			1	02.09.2026	18.07.2027		
33949	42800	BER	LAS			1	02.09.2026	18.07.2027		
33952	42900	BER	LIC			1	02.09.2026	18.07.2027		
33955	43000	BER	MAR			1	02.09.2026	18.07.2027		
33958	43100	BER	MEN			1	02.09.2026	18.07.2027		
33961	43200	BER	MÖL			1	02.09.2026	18.07.2027		
33967	43400	BER	NIE			1	02.09.2026	18.07.2027		
33973	43800	BER	PRN			1	02.09.2026	18.07.2027		
33976	43900	BER	RAM			1	02.09.2026	18.07.2027		
33979	44000	BER	RBG			1	02.09.2026	18.07.2027		
33985	44200	BER	ROS			1	02.09.2026	18.07.2027		
33988	44300	BER	SLF			1	02.09.2026	18.07.2027		
33991	44400	BER	SLL			1	02.09.2026	18.07.2027		
33994	44500	BER	SNL			1	02.09.2026	18.07.2027		
34000	44700	BER	STA			1	02.09.2026	18.07.2027		
34006	46800	BER	ALV			1	02.09.2026	18.07.2027		
34009	57600	BER	GÖB			1	02.09.2026	18.07.2027		
34012	57700	BER	KAN			1	02.09.2026	18.07.2027		
34018	123000	BER	BDN			1	02.09.2026	18.07.2027		
34021	123100	BER	GRU			1	02.09.2026	18.07.2027		
34027	123300	BER	SMI			1	02.09.2026	18.07.2027		
34033	123500	BER	VER			1	02.09.2026	18.07.2027		
34024	123200	BER	MÜN			2	02.09.2026	18.07.2027		
32350	400	BI	RBG	06a		2	02.09.2026	18.07.2027		
32353	5500	BI	RBG	07b		2	02.09.2026	18.07.2027		
32356	52000	BI	HRM	08a		2	02.09.2026	18.07.2027		
32359	54200	BI	SNZ	09c		2	02.09.2026	18.07.2027	532	
32362	55500	BI	SNZ	10b		2	02.09.2026	18.07.2027		
32365	56900	BI	SNZ	10c		2	02.09.2026	18.07.2027	538	
32368	91900	BI	RBG	05a		2	02.09.2026	18.07.2027	538	
32371	93200	BI	EWE	05b		2	02.09.2026	18.07.2027		
32374	94200	BI	SNZ	05c		2	02.09.2026	18.07.2027		
32377	13700	CH	HOM	09c		2	02.09.2026	18.07.2027	537	
32380	19300	CH	HIN	10a		2	02.09.2026	18.07.2027		
32383	21500	CH	HOM	10b		2	02.09.2026	18.07.2027		
32386	22900	CH	HOM	10c		2	02.09.2026	18.07.2027		
32389	52900	CH	VER	08c		3	02.09.2026	18.07.2027	537	
32392	100	D	RBG	06a		3	02.09.2026	18.07.2027	112	
32395	1700	D	EWE	06b		3	02.09.2026	18.07.2027	144	
32398	2600	D	KLU	06c		2	02.09.2026	18.07.2027	114	
32401	3800	D	HIM	07a		3	02.09.2026	18.07.2027	142	
32404	5200	D	KBN	07b		3	02.09.2026	18.07.2027	124	
32407	6300	D	KLU	07c		3	02.09.2026	18.07.2027	134	
32410	7800	D	WOL	08a		2	02.09.2026	18.07.2027	221	
32413	9400	D	MÖL	08b		2	02.09.2026	18.07.2027	122	
32416	10200	D	KBN	08c		3	02.09.2026	18.07.2027	132	
32419	11100	D	KLU	09a		2	02.09.2026	18.07.2027	133	
32422	12400	D	EMT	09b		2	02.09.2026	18.07.2027	223	
32425	13500	D	WOL	09c		2	02.09.2026	18.07.2027	222	
32431	21100	D	WOL	10b		2	02.09.2026	18.07.2027	143	
32434	91600	D	WOL	05a		3	02.09.2026	18.07.2027	147	
32437	92900	D	EWE	05b		4	02.09.2026	18.07.2027	127	
32440	93900	D	KLU	05c		4	02.09.2026	18.07.2027	137	
32443	100400	D	EGL	10a		2	02.09.2026	18.07.2027	113	
32428	19100	D	EGL	10c		2	02.09.2026	18.07.2027	222	
34257	125000	D	EGL	10c	D_10c	1	17.11.2026	06.02.2027		
32446	500	E	BOR	06a		3	02.09.2026	18.07.2027	112	
32449	2100	E	NEM	06b		3	02.09.2026	18.07.2027	144	
32452	3000	E	MAR	06c		3	02.09.2026	18.07.2027	114	
32455	4500	E	FRY	07a		3	02.09.2026	18.07.2027	142	
32458	5800	E	ELS	07b		3	02.09.2026	18.07.2027	124	
32461	6600	E	PRN	07c		3	02.09.2026	18.07.2027	134	
32464	8200	E	NEM	08a		3	02.09.2026	18.07.2027	221	
32467	10100	E	BOR	08b		3	02.09.2026	18.07.2027	122	
32476	12900	E	ELS	09b		2	02.09.2026	18.07.2027	223	
32479	14900	E	HOM	09c		2	02.09.2026	18.07.2027	222	
32485	21600	E	HIM	10b		3	02.09.2026	18.07.2027		
32488	22500	E	NIE	10c		2	02.09.2026	18.07.2027	123	
32491	92000	E	ELS	05a		3	02.09.2026	18.07.2027	147	
32494	93300	E	HIM	05b		4	02.09.2026	18.07.2027	127	
32497	94300	E	PRN	05c		3	02.09.2026	18.07.2027	137	
32470	10700	E	MÜN	08c		1	02.09.2026	18.07.2027	132	
32473	11700	E	MAR	09a		2	02.09.2026	18.07.2027	133	
34221	124100	E	BOR	08c	E_08c	1	02.09.2026	18.07.2027	132	
32482	19500	E	PRN	10a		2	02.09.2026	18.07.2027	113	
34260	124800	E	PRN	10a	E_10a	1	17.11.2026	06.02.2027	132	
32503	9700	EK	BDN	08b		2	02.09.2026	18.07.2027	122	
32512	52800	EK	STA	08c		2	02.09.2026	18.07.2027	132	
32515	53200	EK	STA	09a		2	02.09.2026	18.07.2027	133	
32521	94900	EK	STA	06a		3	02.09.2026	18.07.2027	112	
32506	12500	EK	WES	09b		2	02.09.2026	18.07.2027	223	
32509	51500	EK	MÜN	07c		1	02.09.2026	18.07.2027	134	
32518	54000	EK	MÜN	09c		2	02.09.2026	18.07.2027	222	
34224	123900	EK	STR	07c	EK_07c	1	02.09.2026	18.07.2027	134	
34246	124300	ER	KBN	08a		2	02.09.2026	18.07.2027	221	
32527	7700	GE	KLU	07c		2	02.09.2026	18.07.2027	134	
32530	9500	GE	EGL	08b		3	02.09.2026	18.07.2027	122	
32536	13600	GE	EGL	09c		2	02.09.2026	18.07.2027	222	
32539	37000	GE	LIC	10a		2	02.09.2026	18.07.2027	113	
32542	50800	GE	KLU	07a		2	02.09.2026	18.07.2027	142	
32545	95400	GE	KRA	06b		3	02.09.2026	18.07.2027	144	
32548	95600	GE	LIC	06c		3	02.09.2026	18.07.2027	114	
32524	5300	GE	VAS	07b		2	02.09.2026	18.07.2027	124	
32533	10300	GE	SNL	08c		3	02.09.2026	15.11.2026	132	
32533	10300	GE	VAS	08c		3	16.11.2026	18.07.2027	132	
32551	49900	IF	KLN	06a		1	02.09.2026	18.07.2027	PC	
32554	50400	IF	KLN	06b		1	02.09.2026	18.07.2027	PC	
32557	50500	IF	WOL	06c		1	02.09.2026	18.07.2027	PC	
32566	600	KU	WIS	06a		2	02.09.2026	18.07.2027	KU2	
32575	13000	KU	EMT	09b		2	02.09.2026	18.07.2027	KU2	
32578	23000	KU	EMT	10c		2	02.09.2026	18.07.2027	KU2	
32581	53300	KU	WIS	09a		2	02.09.2026	18.07.2027	KU2	
32584	92100	KU	STA	05a		1	02.09.2026	18.07.2027	KU2	
32569	3100	KU	MÜN	06c		2	02.09.2026	18.07.2027	KU2	
32572	10800	KU	AHL	08c		1	02.09.2026	18.07.2027	KU1	
32587	93400	KU	MÜN	05b		1	02.09.2026	18.07.2027	KU1	
32590	94400	KU	MÜN	05c		1	02.09.2026	18.07.2027	KU2	
32596	107900	KU	MÜN	10a		2	02.09.2026	18.07.2027	KU1	
34227	123800	KU	STA	08c	KU_08c	1	02.09.2026	18.07.2027	KU1	
32599	300	M	GRE	06a		3	02.09.2026	18.07.2027	112	
32602	1900	M	KLN	06b		3	02.09.2026	18.07.2027	144	
32605	2800	M	HIN	06c		3	02.09.2026	18.07.2027	114	
32608	4000	M	HIN	07a		3	02.09.2026	18.07.2027	142	
32614	6500	M	SNL	07c		3	02.09.2026	18.07.2027	134	
32617	9800	M	GRU	08b		3	02.09.2026	18.07.2027	122	
32620	10400	M	SLL	08c		3	02.09.2026	18.07.2027	132	
32623	11300	M	MEN	09a		2	02.09.2026	18.07.2027	133	
32626	12700	M	SNL	09b		2	02.09.2026	18.07.2027	223	
32632	14500	M	MAR	08a		3	02.09.2026	18.07.2027	221	
32638	22800	M	SNL	10c		2	02.09.2026	18.07.2027	123	
32641	91800	M	GRU	05a		4	02.09.2026	18.07.2027	147	
32647	94100	M	HIN	05c		4	02.09.2026	18.07.2027	137	
32650	99400	M	GRE	10a		3	02.09.2026	18.07.2027	113	
32611	5400	M	HIN	07b		3	02.09.2026	18.07.2027	124	
32629	13800	M	WEB	09c		2	02.09.2026	18.07.2027	222	
32644	93100	M	KLN	05b		3	02.09.2026	18.07.2027		
32635	19200	M	RIC	10b		2	02.09.2026	18.07.2027	143	
34263	124900	M	RIC	10b	M_10b	1	17.11.2026	06.02.2027		
32653	800	MU	WIL	06a		1	02.09.2026	18.07.2027	MU1	
32656	3200	MU	WÖR	06c		1	02.09.2026	18.07.2027	MU1	
32659	3300	MU	WÖR	06b		1	02.09.2026	18.07.2027	MU1	
32662	6900	MU	WÖR	07c		2	02.09.2026	18.07.2027	MU1	
32665	8800	MU	WIL	08a		2	02.09.2026	18.07.2027		
32668	36900	MU	WÖR	10b		2	02.09.2026	18.07.2027	MU1	
32671	52700	MU	MÖL	08b		2	02.09.2026	18.07.2027		
32674	53700	MU	WIL	09b		2	02.09.2026	18.07.2027	MU1	
32677	55900	MU	WÖR	10c		2	02.09.2026	18.07.2027	MU1	
32680	92300	MU	WÖR	05a		1	02.09.2026	18.07.2027	MU1	
32683	93500	MU	GRU	05b		1	02.09.2026	18.07.2027	MU2	
32686	94500	MU	WÖR	05c		1	02.09.2026	18.07.2027	MU1	
32689	37700	ORC	MÖL	AG		1	02.09.2026	18.07.2027	AUL	
32692	92700	ORI	GUS	05a		1	02.09.2026	18.07.2027	147	
32695	93700	ORI	EWE	05b		1	02.09.2026	18.07.2027	127	
32698	94700	ORI	PRN	05c		1	02.09.2026	18.07.2027	137	
32701	97000	ORI	WES	08a		1	02.09.2026	18.07.2027	221	
32704	97200	ORI	MÖL	08b		1	02.09.2026	18.07.2027	122	
32707	97400	ORI	SLL	08c		1	02.09.2026	18.07.2027	132	
32710	5600	PH	VER	07b		2	02.09.2026	18.07.2027	PH1	
32713	11500	PH	GRE	09a		2	02.09.2026	18.07.2027		
32716	21300	PH	SLL	10b		2	02.09.2026	18.07.2027		
32719	50900	PH	WEB	07a		2	02.09.2026	18.07.2027	PH1	
32722	51600	PH	GRE	07c		2	02.09.2026	18.07.2027		
32725	53600	PH	WEB	09b		2	02.09.2026	18.07.2027	PH2	
32728	54600	PH	SLL	10a		2	02.09.2026	18.07.2027		
32731	95500	PH	VER	06b		3	02.09.2026	18.07.2027		
32734	95700	PH	VER	06c		3	02.09.2026	18.07.2027	PH1	
32743	91400	SG	TÖN	05c	SG_05c	2	02.09.2026	18.07.2027	137	
32746	91500	SG	TÖN	06c		2	02.09.2026	18.07.2027	AUL	
32740	91300	SG	TÖN	AG	SG_AG	1	02.09.2026	18.07.2027	AUL	
32749	700	SP	GÖB	06a		2	02.09.2026	18.07.2027		
32755	3500	SP	LAS	06c		2	02.09.2026	18.07.2027		
32758	5100	SP	FRY	07a		2	02.09.2026	18.07.2027		
32761	6000	SP	KAN	07b		2	02.09.2026	18.07.2027	GR_SPH	
32764	7000	SP	LAS	07c		2	02.09.2026	18.07.2027		
32767	9300	SP	WES	08a		3	02.09.2026	18.07.2027		
32770	10000	SP	FRY	08b		2	02.09.2026	18.07.2027		
32773	12300	SP	FRY	09a		2	02.09.2026	18.07.2027		
32776	13100	SP	GÖB	09b		2	02.09.2026	18.07.2027	KL_SPH	
32779	14100	SP	WES	09c		2	02.09.2026	18.07.2027		
32782	14800	SP	GÖB	08c		2	02.09.2026	18.07.2027		
32785	21000	SP	XHO	10a		2	02.09.2026	18.07.2027		
32788	22400	SP	KAN	10b		2	02.09.2026	18.07.2027		
32791	23200	SP	GÖB	10c		2	02.09.2026	18.07.2027	KL_SPH	
32794	92200	SP	GUS	05a		3	02.09.2026	18.07.2027		
32800	94600	SP	GÖB	05c		3	02.09.2026	18.07.2027		
32752	2300	SP	GÖB	06b		3	02.09.2026	18.07.2027	GR_SPH	
32797	93600	SP	WES	05b		3	02.09.2026	18.07.2027	GR_SPH	
32803	37400	UCO	MÖL	AG		1	02.09.2026	18.07.2027	AUL	
32812	66800	F71	ALV	09a	F71_09	3	02.09.2026	18.07.2027	133	
32818	95900	F71	BDN	07a~07b	F71_07	3	02.09.2026	18.07.2027	142	
32815	70800	F71	KAN	10a~10b	F71_10	3	02.09.2026	18.07.2027	113	
32809	65700	F71	WIS	08a	F71_08	3	02.09.2026	18.07.2027	221	
32824	96000	F72	KAN	07b~07c	F72_07	3	02.09.2026	18.07.2027	124	
32830	97500	F72	BDN	09b	F72_09	3	02.09.2026	18.07.2027	223	
32833	98800	F72	BDN	10a~10c	F72_10	3	02.09.2026	18.07.2027	143	
34249	124400	F72	HRM	08b	F72_08	3	02.09.2026	18.07.2027	122	
32839	97600	F73	HRM	09c	F73_09	3	02.09.2026	18.07.2027	222	
34252	124500	F73	ALV	08c	F73_08	3	02.09.2026	18.07.2027	123	
32842	60700	SWS1	KLU	06c	SWS1_06c	1	02.09.2026	18.07.2027	114	
32845	73700	SWS1	EWE	06b	SWS1_06b	1	02.09.2026	18.07.2027	144	
32848	74300	SWS1	RBG	06a	SWS1_06a	1	02.09.2026	18.07.2027	112	
32851	68900	SWS2	KLU	06b	SWS2_06b	1	02.09.2026	18.07.2027	143	
32854	72600	SWS2	EWE	06c	SWS2_06c	1	02.09.2026	18.07.2027		
32857	73400	SWS2	WOL	06a	SWS2_06a	1	02.09.2026	18.07.2027	133	
32869	65600	ER1	PRN	10a	ER1_10	2	02.09.2026	18.07.2027	113	
32872	92400	ER1	ELS	05a~05b	ER1_05	2	02.09.2026	18.07.2027	147	
32875	95000	ER1	BOR	06a~06b	ER1_06	2	02.09.2026	18.07.2027	112	
32884	97800	ER1	KBN	09a	ER1_09	2	02.09.2026	18.07.2027	133	
32878	96300	ER1	ELS	07a~07c	ER1_07	2	02.09.2026	18.07.2027	142	
32887	92500	ER2	PRN	05b~05c	ER2_05	2	02.09.2026	18.07.2027	127	
32890	95100	ER2	KBN	06b~06c	ER2_06	2	02.09.2026	18.07.2027	144	
32896	97900	ER2	PJU	09b	ER2_09	2	02.09.2026	18.07.2027	223	
32899	99500	ER2	PJU	10b	ER2_10	2	02.09.2026	18.07.2027	143	
32893	96400	ER2	KBN	07a~07b	ER2_07	2	02.09.2026	18.07.2027	124	
33494	116900	ER2	SLF	EF	ER2_EF	2	02.09.2026	18.07.2027		
32905	98000	ER3	ELS	09c	ER3_09	2	02.09.2026	18.07.2027	222	
32908	99600	ER3	SLF	10c	ER3_10	2	02.09.2026	18.07.2027	123	
32911	98400	GEW1	WES	09a~09b~09c	GEW1_09	3	02.09.2026	18.07.2027	223	
32914	100000	GEW1	STA	10a~10b~10c	GEW1_10	2	02.09.2026	18.07.2027	123	
33453	108900	IF1	WOL	09a~09b~09c	IF1_09	3	02.09.2026	18.07.2027	PC	
33456	109000	IF1	WOL	10a~10b	IF1_10	2	02.09.2026	18.07.2027	PC	
32923	98100	KR1	ALV	09a~09b~09c	KR1_09	2	02.09.2026	18.07.2027	513	
32926	99700	KR1	NIE	10a~10b~10c	KR1_10	2	02.09.2026	18.07.2027	522	
32929	100500	KR1	ALV	05a~05b~05c	KR1_05	2	02.09.2026	18.07.2027	137	
32920	96500	KR1	LIC	07b~07c	KR1_07	2	02.09.2026	18.07.2027		
32917	60000	KR1	LIC	06a	KR1_06	2	02.09.2026	18.07.2027	114	
32935	95200	KR2	NIE	06b~06c	KR2_06	2	02.09.2026	18.07.2027	523	
32938	96600	KR2	NIE	07a~07b	KR2_07	2	02.09.2026	18.07.2027		
32947	96100	L71	ROS	07a~07b	L71_07	3	02.09.2026	18.07.2027	134	
32950	96800	L71	ROS	08a~08c	L71_08	3	02.09.2026	18.07.2027	132	
32953	97700	L71	ROS	09a~09b~09c	L71_09	3	02.09.2026	18.07.2027		
32944	66600	L71	GÖB	10a~10b	L71_10	3	02.09.2026	18.07.2027		
32959	67900	L72	ROS	08b	L72_08	3	02.09.2026	18.07.2027		
32962	96200	L72	PJU	07b~07c	L72_07	3	02.09.2026	18.07.2027	143	
32965	98900	L72	MEN	10b~10c	L72_10	3	02.09.2026	18.07.2027		
33459	108800	NW1	LAS	09a~09b~09c	NW1_09	3	02.09.2026	18.07.2027	533	
34255	124200	NW1	RIC	10a~10b~10c	NW1_10	2	02.09.2026	18.07.2027		
34266	124600	NWG1	HIN	06a	NWG1_06a	1	02.09.2026	18.07.2027	537	
34269	124700	NWG2	KLN	06b	NWG2_06b	1	02.09.2026	18.07.2027	538	
32980	98700	TWM1	RBG	09a~09b~09c	TWM1_09	3	02.09.2026	18.07.2027	124	
32983	100200	TWM1	WÖR	10a~10b~10c	TWM1_10	2	02.09.2026	18.07.2027	AUL	
32986	98200	S91	ALV	09a~09b~09c	S91_09	3	02.09.2026	18.07.2027	133	
32989	99800	S91	RAM	10a~10b~10c	S91_10	3	02.09.2026	18.07.2027		
33497	115700	S91	RAM	EF	S91_EF	2	02.09.2026	18.07.2027		
33500	110700	MU_G2	MÖL	Q2	MU_G2_Q2	2	02.09.2026	18.07.2027	MU1	
33503	112300	BI_G1	EWE	Q2	BI_G1_Q2	2	02.09.2026	18.07.2027	532	
33506	115000	BI_G1	HRM	EF	BI_G1_EF	2	02.09.2026	18.07.2027	532	
33509	120500	BI_G1	EWE	Q1	BI_G1_Q1	2	02.09.2026	18.07.2027	538	
33512	110900	BI_G2	LAS	Q2	BI_G2_Q2	2	02.09.2026	18.07.2027	532	
33515	115800	BI_G2	HRM	EF	BI_G2_EF	2	02.09.2026	18.07.2027		
33518	120900	BI_G2	SNZ	Q1	BI_G2_Q1	2	02.09.2026	18.07.2027	532	
33521	115400	BI_G3	SNZ	EF	BI_G3_EF	2	02.09.2026	18.07.2027		
33524	109300	BI_L1	RBG	Q2	BI_L1_Q2	4	02.09.2026	18.07.2027	538	
33527	118600	BI_L1	LAS	Q1	BI_L1_Q1	4	02.09.2026	18.07.2027		
33530	79600	CH_G1	VER	Q1	CH_G1_Q1	2	02.09.2026	18.07.2027	537	
33533	112600	CH_G1	KLN	Q2	CH_G1_Q2	2	02.09.2026	18.07.2027		
33536	115900	CH_G1	VER	EF	CH_G1_EF	2	02.09.2026	18.07.2027	537	
33539	116500	CH_G2	VER	EF	CH_G2_EF	2	02.09.2026	18.07.2027		
33542	113200	D_G1	KBN	Q2	D_G1_Q2	2	02.09.2026	18.07.2027		
33545	117600	D_G1	EMT	EF	D_G1_EF	2	02.09.2026	18.07.2027	516	
33548	121000	D_G1	KBN	Q1	D_G1_Q1	2	02.09.2026	18.07.2027		
33551	110400	D_G2	SLF	Q2	D_G2_Q2	2	02.09.2026	18.07.2027	516	
33554	114300	D_G2	KAT	EF	D_G2_EF	2	02.09.2026	18.07.2027		
33557	120000	D_G2	RBG	Q1	D_G2_Q1	2	02.09.2026	18.07.2027	513	
33560	117200	D_G3	SLF	EF	D_G3_EF	3	02.09.2026	18.07.2027		
33563	119200	D_G3	WOL	Q1	D_G3_Q1	2	02.09.2026	18.07.2027		
33566	109800	D_L1	BOR	Q2	D_L1_Q2	4	02.09.2026	18.07.2027		
33569	118700	D_L1	SLF	Q1	D_L1_Q1	4	02.09.2026	18.07.2027		
33572	113300	E_G1	ELS	Q2	E_G1_Q2	2	02.09.2026	18.07.2027		
33575	114000	E_G1	ELS	EF	E_G1_EF	2	02.09.2026	18.07.2027	514	
33577	120300	E_G1	HOM	Q1	E_G1_Q1	2	02.09.2026	18.07.2027		
33580	111000	E_G2	ELS	Q2	E_G2_Q2	2	02.09.2026	18.07.2027		
33583	114700	E_G2	MAR	EF	E_G2_EF	2	02.09.2026	18.07.2027		
33586	119300	E_G2	NIE	Q1	E_G2_Q1	2	02.09.2026	18.07.2027	514	
33589	112900	E_G3	NIE	Q2	E_G3_Q2	2	02.09.2026	18.07.2027	514	
33592	113800	E_G3	VAS	EF	E_G3_EF	1	02.09.2026	18.07.2027	521	
34272	81500	E_G3	ELS	EF	E_G3_EF_1	1	02.09.2026	18.07.2027	521	
33595	109500	E_L1	HIM	Q2	E_L1_Q2	4	02.09.2026	18.07.2027		
33598	118800	E_L1	MAR	Q1	E_L1_Q1	4	02.09.2026	18.07.2027	514	
33601	109900	E_L2	NEM	Q2	E_L2_Q2	4	02.09.2026	18.07.2027		
33604	118200	E_L2	NEM	Q1	E_L2_Q1	4	02.09.2026	18.07.2027		
33607	111500	EK_G1	KRA	Q2	EK_G1_Q2	3	02.09.2026	18.07.2027	525	
33613	119500	EK_G1	WES	Q1	EK_G1_Q1	2	02.09.2026	18.07.2027		
33610	114400	EK_G1	STR	EF	EK_G1_EF	1	02.09.2026	18.07.2027	525	
34275	125200	EK_G1	MÜN	EF	EK_G1_EF_1	1	02.09.2026	18.07.2027	526	
33616	114100	EK_G2	MÜN	EF	EK_G2_EF	1	02.09.2026	18.07.2027	525	
33619	119700	EK_G2	MÜN	Q1	EK_G2_Q1	2	02.09.2026	18.07.2027	523	
34278	125300	EK_G2	STR	EF	EK_G2_EF_1	1	02.09.2026	18.07.2027	525	
33622	110000	EK_L1	KRA	Q2	EK_L1_Q2	4	02.09.2026	18.07.2027	525	
33625	118300	EK_L1	BDN	Q1	EK_L1_Q1	4	02.09.2026	18.07.2027	525	
33628	79100	ER_G1	PJU	Q1	ER_G1_Q1	2	02.09.2026	18.07.2027	517	
33631	116200	ER_G1	KBN	EF	ER_G1_EF	2	02.09.2026	18.07.2027	523	
33634	121600	ER_G1	PJU	Q2	ER_G1_Q2	2	02.09.2026	18.07.2027	517	
33637	112400	ER_G2	PJU	Q2	ER_G2_Q2	2	02.09.2026	18.07.2027	517	
33640	119600	ER_G2	PRN	Q1	ER_G2_Q1	2	02.09.2026	18.07.2027		
33643	78500	ER_G3	SLF	Q1	ER_G3_Q1	2	02.09.2026	18.07.2027	517	
33646	110500	F7_G1	KAN	Q2	F7_G1_Q2	2	02.09.2026	18.07.2027	521	
33649	114500	F7_G1	BDN	EF	F7_G1_EF	2	02.09.2026	18.07.2027	521	
33652	120100	F7_G1	KAN	Q1	F7_G1_Q1	2	02.09.2026	18.07.2027		
33655	111900	GE_G1	LIC	Q2	GE_G1_Q2	2	02.09.2026	18.07.2027		
33658	115500	GE_G1	EGL	EF	GE_G1_EF	2	02.09.2026	18.07.2027	528	
33661	119800	GE_G1	EGL	Q1	GE_G1_Q1	3	02.09.2026	18.07.2027	528	
33664	88100	GE_G2	SMI	Q1	GE_G2_Q1	2	02.09.2026	18.07.2027	511	
33667	113400	GE_G2	NEM	Q2	GE_G2_Q2	2	02.09.2026	18.07.2027	528	
33670	116000	GE_G2	EGL	EF	GE_G2_EF	2	02.09.2026	18.07.2027		
33673	110100	GE_L1	EGL	Q2	GE_L1_Q2	4	02.09.2026	18.07.2027		
33676	118400	GE_L1	ROS	Q1	GE_L1_Q1	4	02.09.2026	18.07.2027	528	
33679	112500	IF_G1	MEN	Q2	IF_G1_Q2	2	02.09.2026	18.07.2027	PC	
33682	115600	IF_G1	MEN	EF	IF_G1_EF	2	02.09.2026	18.07.2027	PC	
33685	120200	IF_G1	WOL	Q1	IF_G1_Q1	2	02.09.2026	18.07.2027	PC	
33688	79300	KR_G1	LIC	Q1	KR_G1_Q1	2	02.09.2026	18.07.2027		
33691	117000	KR_G1	LIC	EF	KR_G1_EF	2	02.09.2026	18.07.2027	526	
33694	121700	KR_G1	NIE	Q2	KR_G1_Q2	2	02.09.2026	18.07.2027		
33697	78100	KU_G1	EMT	Q1	KU_G1_Q1	2	02.09.2026	18.07.2027		
33703	121800	KU_G1	WIS	Q2	KU_G1_Q2	2	02.09.2026	18.07.2027	KU1	
33700	87300	KU_G1	AHL	EF	KU_G1_EF	1	02.09.2026	18.07.2027	KU2	
34281	87400	KU_G1	WIS	EF	KU_G1_EF_1	1	02.09.2026	18.07.2027	KU1	
33706	111700	KU_G2	WIS	Q2	KU_G2_Q2	3	02.09.2026	18.07.2027		
33709	117300	KU_G2	EMT	EF	KU_G2_EF	3	02.09.2026	18.07.2027		
33712	120600	KU_G2	EMT	Q1	KU_G2_Q1	2	02.09.2026	18.07.2027	KU2	
33715	88500	L7_G1	MEN	Q1~Q2	L7_G1_Q1_Q2	2	02.09.2026	18.07.2027	523	
33718	117700	L7_G1	PJU	EF	L7_G1_EF	2	02.09.2026	18.07.2027		
33721	110600	M_G1	GRE	Q2	M_G1_Q2	2	02.09.2026	18.07.2027		
33724	114600	M_G1	GRU	EF	M_G1_EF	2	02.09.2026	18.07.2027		
33727	120400	M_G1	GRE	Q1	M_G1_Q1	2	02.09.2026	18.07.2027		
33730	111100	M_G2	KLN	Q2	M_G2_Q2	2	02.09.2026	18.07.2027		
33733	116300	M_G2	GRU	EF	M_G2_EF	2	02.09.2026	18.07.2027		
33736	119900	M_G2	WEB	Q1	M_G2_Q1	3	02.09.2026	18.07.2027	515	
33739	112100	M_G3	WEB	Q2	M_G3_Q2	2	02.09.2026	18.07.2027		
33742	114200	M_G3	HIN	EF	M_G3_EF	2	02.09.2026	18.07.2027	515	
33745	116600	M_G4	MAR	EF	M_G4_EF	2	02.09.2026	18.07.2027		
33748	76700	M_L1	GRE	Q1	M_L1_Q1	4	02.09.2026	18.07.2027	515	
33751	110300	M_L1	MEN	Q2	M_L1_Q2	4	02.09.2026	18.07.2027		
33754	109600	M_L2	WEB	Q2	M_L2_Q2	4	02.09.2026	18.07.2027		
33757	118900	M_L2	SNL	Q1	M_L2_Q1	4	02.09.2026	18.07.2027		
33763	113900	MU_G1	WIL	EF	MU_G1_EF	2	02.09.2026	18.07.2027	MU2	
33766	122000	MU_G1	GRU	Q2	MU_G1_Q2	2	02.09.2026	18.07.2027	MU2	
33760	77700	MU_G1	GRU	Q1	MU_G1_Q1	2	02.09.2026	18.07.2027		
33769	111200	PH_G1	WEB	Q2	PH_G1_Q2	2	02.09.2026	18.07.2027	PH2	
33772	116800	PH_G1	GRE	EF	PH_G1_EF	2	02.09.2026	18.07.2027	PH1	
33775	121100	PH_G1	VER	Q1	PH_G1_Q1	2	02.09.2026	18.07.2027	PH2	
33778	117400	PH_G2	WEB	EF	PH_G2_EF	3	02.09.2026	18.07.2027		
33781	109700	PH_L1	RIC	Q2	PH_L1_Q2	4	02.09.2026	18.07.2027		
33784	119000	PH_L1	SLL	Q1	PH_L1_Q1	4	02.09.2026	18.07.2027		
33787	111800	PL_G1	KAT	Q2	PL_G1_Q2	3	02.09.2026	18.07.2027		
33790	116700	PL_G1	KAT	EF	PL_G1_EF	2	02.09.2026	18.07.2027	521	
33793	120700	PL_G1	KAT	Q1	PL_G1_Q1	2	02.09.2026	18.07.2027		
33796	113000	S9_G1	ALV	Q2	S9_G1_Q2	2	02.09.2026	18.07.2027	513	
33799	120800	S9_G1	RAM	Q1	S9_G1_Q1	2	02.09.2026	18.07.2027	522	
33802	119100	SP_G2	LAS	Q1	SP_G2_Q1	2	02.09.2026	18.07.2027		
33805	111300	SP1_G1	FRY	Q2	SP1_G1_Q2	2	02.09.2026	18.07.2027	GR_SPH	
33808	117800	SP1_G1	KAN	EF	SP1_G1_EF	2	02.09.2026	18.07.2027	SB	
33811	121300	SP1_G1	GÖB	Q1	SP1_G1_Q1	2	02.09.2026	18.07.2027		
33814	113100	SP2_G1	GÖB	Q2	SP2_G1_Q2	2	02.09.2026	18.07.2027		
33817	75700	SP2_G2	XHO	Q1	SP2_G2_Q1	2	02.09.2026	18.07.2027	KL_SPH	
33820	122100	SP2_G2	LAS	Q2	SP2_G2_Q2	2	02.09.2026	18.07.2027	KL_SPH	
33823	111400	SP3_G1	GUS	Q2	SP3_G1_Q2	2	02.09.2026	18.07.2027	GR_SPH	
33826	117100	SP3_G1	KAN	EF	SP3_G1_EF	2	02.09.2026	18.07.2027		
33829	88700	SW_G1	SMI	Q1	SW_G1_Q1	3	02.09.2026	18.07.2027	511	
33832	114800	SW_G1	KRA	EF	SW_G1_EF	2	02.09.2026	18.07.2027	523	
33835	121400	SW_G1	KRA	Q2	SW_G1_Q2	2	02.09.2026	18.07.2027		
33838	78200	SW_G2	SMI	Q1	SW_G2_Q1	2	02.09.2026	18.07.2027		
33841	115200	SW_G2	SMI	EF	SW_G2_EF	2	02.09.2026	18.07.2027		
33844	116400	SW_G3	SMI	EF	SW_G3_EF	2	02.09.2026	18.07.2027	524	
33847	110200	SW_L1	GUS	Q2	SW_L1_Q2	4	02.09.2026	18.07.2027		
33850	109400	D_L2	KRS	Q2	D_L2_Q2	4	02.09.2026	18.07.2027	516	
33853	115100	EK_G3	STR	EF	EK_G3_EF	1	02.09.2026	18.07.2027	513	
34284	86100	EK_G3	MÜN	EF	EK_G3_EF_1	1	02.09.2026	18.07.2027	525	
33355	36400	ESL	ALV	AG		1	02.09.2026	18.07.2027		
33358	91000	LR	BDN	AG		1	02.09.2026	18.07.2027	WR	
33361	57000	MAV	HIN	AG		1	02.09.2026	18.07.2027		
33856	117900	M_V1	XHO	EF	M_V1_EF	2	02.09.2026	18.07.2027		
33367	92800	ITG	GUS	05a		1	02.09.2026	18.07.2027	PC	
33370	93800	ITG	KLN	05b		1	02.09.2026	18.07.2027	PC	
33373	94800	ITG	PRN	05c		1	02.09.2026	18.07.2027	PC	
33376	8100	WP	SMI	08a		3	02.09.2026	18.07.2027	221	
33379	9600	WP	LIC	08b		3	02.09.2026	18.07.2027	122	
33382	11200	WP	KRA	09a		2	02.09.2026	18.07.2027	133	
33385	91700	WP	GUS	05a		2	02.09.2026	18.07.2027	147	
33388	93000	WP	SMI	05b		1	02.09.2026	18.07.2027	127	
33391	94000	WP	PRN	05c		1	02.09.2026	18.07.2027	137	
33394	95800	WP	SMI	07a		2	02.09.2026	18.07.2027	142	
33859	78000	IV_G1	STR	Q1	IV_G1_Q1	1	02.09.2026	18.07.2027	MU2	
33862	116100	KU_G3	AHL	EF	KU_G3_EF	1	02.09.2026	18.07.2027	KU2	
34287	125100	KU_G3	WIS	EF	KU_G3_EF_1	1	02.09.2026	18.07.2027		
34042	37200	BRG	EGL	AG		2	02.09.2026	18.07.2027		
33415	90800	DaZ	GRU	06b~08b~09c		1	02.09.2026	18.07.2027	144	
33865	112200	GE_Z1	VAS	Q2	GE_Z1_Q2	2	02.09.2026	18.07.2027		
33868	110800	GE_Z2	VAS	Q2	GE_Z2_Q2	2	02.09.2026	18.07.2027		
33871	114900	SP1_G2	KAN	EF	SP1_G2_EF	2	02.09.2026	18.07.2027		
33874	121500	SW_Z1	SMI	Q2	SW_Z1_Q2	2	02.09.2026	18.07.2027		
33877	112000	SW_Z2	KRA	Q2	SW_Z2_Q2	2	02.09.2026	18.07.2027	513	
33880	112800	S0_G1	RAM	Q2	S0_G1_Q2	3	02.09.2026	18.07.2027	522	
33883	115300	S0_G1	AHL	EF	S0_G1_EF	3	02.09.2026	18.07.2027		
33886	119400	S0_G1	RAM	Q1	S0_G1_Q1	3	02.09.2026	18.07.2027	522	
33887	117500	S0_G2	AHL	EF	S0_G2_EF	3	02.09.2026	18.07.2027	522	
33890	111600	ER_P1	PJU	Q2	ER_P1_Q2	3	02.09.2026	18.07.2027	517	
33893	118500	KU_L1	EMT	Q1	KU_L1_Q1	4	02.09.2026	18.07.2027		
33896	121900	MU_P1	MÖL	Q2	MU_P1_Q2	2	02.09.2026	18.07.2027		
33899	122300	E_V1	HOM	EF	E_G4_EF	2	02.09.2026	18.07.2027		
`;

function seedAusUntisCsv(csvText, klassen) {
  const SEK1_MUSTER = /^(0[5-9]|10)[a-c]$/;
  const zeilenRoh = csvText.trim().split(/\r?\n/);
  if (zeilenRoh.length < 2) return { faecher: [], lehrer: [], lerngruppen: [] };
  const header = zeilenRoh[0].split("\t");
  const rows = zeilenRoh.slice(1).map((zeile) => {
    const werte = zeile.split("\t");
    const obj = {};
    header.forEach((h, i) => (obj[h] = werte[i] || ""));
    return obj;
  });

  const zeilen = rows
    .map((r) => ({ subject: (r.subject || "").trim(), teacher: (r.teacher || "").trim(), klassen: (r.klassen || "").trim() }))
    .filter((r) => r.subject && r.teacher && r.klassen && r.klassen.split("~").every((k) => SEK1_MUSTER.test(k.trim())));

  const faecher = [];
  const lehrer = [];
  const lerngruppen = [];
  const findeKlasse = (token) => {
    const jahrgang = Number(token.slice(0, 2));
    const buchstabe = token.slice(2);
    return klassen.find((k) => k.jahrgang === jahrgang && k.buchstabe === buchstabe);
  };
  const findeOderErstelleFach = (code) => {
    let f = faecher.find((x) => x.kuerzel === code);
    if (!f) {
      f = { id: uid("f"), name: code, kuerzel: code, quelle: "untis" };
      faecher.push(f);
    }
    return f;
  };
  const findeOderErstelleLehrer = (kuerzel) => {
    let l = lehrer.find((x) => x.kuerzel === kuerzel);
    if (!l) {
      l = { id: uid("l"), name: "", kuerzel, email: "", quelle: "untis" };
      lehrer.push(l);
    }
    return l;
  };

  const gesehen = new Set();
  zeilen.forEach((r) => {
    const klassenTokens = r.klassen.split("~").map((t) => t.trim());
    const klassenObjekte = klassenTokens.map(findeKlasse).filter(Boolean);
    if (klassenObjekte.length === 0) return;
    const klassenIds = klassenObjekte.map((k) => k.id);
    const schluessel = r.subject + "|" + [...klassenIds].sort().join(",");
    if (gesehen.has(schluessel)) return;
    gesehen.add(schluessel);
    const fachObj = findeOderErstelleFach(r.subject);
    const lehrerObj = findeOderErstelleLehrer(r.teacher);
    lerngruppen.push({
      id: uid("lg"),
      fachId: fachObj.id,
      bezeichnung: klassenTokens.join("/"),
      jahrgang: klassenObjekte[0].jahrgang,
      lehrerId: lehrerObj.id,
      klassenIds,
      quelle: "untis",
    });
  });

  return { faecher, lehrer, lerngruppen };
}

function seedInitialerZustand() {
  const klassen = seedKlassen([]);
  const { faecher, lehrer, lerngruppen } = seedAusUntisCsv(STANDARD_UNTIS_EXPORT, klassen);
  return { klassen, faecher, lehrer, lerngruppen };
}

function seedMethoden() {
  return [
    {
      id: uid("m"),
      name: "Placemat",
      beschreibung: "Kooperative Methode zur Ideensammlung in Kleingruppen.",
      faecherIds: [],
      jahrgaenge: [5, 6],
      halbjahr: 1,
      materialien: ["Placemat_Vorlage.pdf"],
    },
    {
      id: uid("m"),
      name: "Lerntempoduett",
      beschreibung: "Schüler bearbeiten Aufgaben im eigenen Tempo und tauschen sich paarweise aus.",
      faecherIds: [],
      jahrgaenge: [6, 7],
      halbjahr: 2,
      materialien: [],
    },
    {
      id: uid("m"),
      name: "Kugellager",
      beschreibung: "Doppelkreis-Methode zum strukturierten Partneraustausch. <strong>Wichtig:</strong> ausreichend Platz für zwei konzentrische Stuhlkreise einplanen. <ul><li>Innenkreis bleibt sitzen</li><li>Außenkreis rotiert im Takt</li></ul>",
      faecherIds: [],
      jahrgaenge: [7, 8],
      halbjahr: 1,
      materialien: ["Kugellager_Anleitung.pdf"],
    },
    {
      id: uid("m"),
      name: "Concept Map",
      beschreibung: "Visuelle Strukturierung von Zusammenhängen im Fachinhalt.",
      faecherIds: [],
      jahrgaenge: [8, 9],
      halbjahr: 2,
      materialien: [],
    },
    {
      id: uid("m"),
      name: "Gruppenpuzzle",
      beschreibung: "Arbeitsteilige Erarbeitung in Expertengruppen mit anschließendem Austausch.",
      faecherIds: [],
      jahrgaenge: [8, 9],
      halbjahr: 2,
      materialien: ["Gruppenpuzzle_Expertentexte.pdf"],
    },
    {
      id: uid("m"),
      name: "Mindmap erstellen",
      beschreibung: "Strukturierung von Fachinhalten in einer visuellen Übersicht. <em>Eignet sich besonders zur Wiederholung vor Klassenarbeiten.</em>",
      faecherIds: [],
      jahrgaenge: [7, 8],
      halbjahr: 1,
      materialien: ["Mindmap_Beispiel.pdf"],
    },
    {
      id: uid("m"),
      name: "Textmarkierung",
      beschreibung: "Systematisches Kennzeichnen von Kernaussagen in Sachtexten.",
      faecherIds: [],
      jahrgaenge: [8, 9],
      halbjahr: 2,
      materialien: ["Textmarkierung_Kurzanleitung.pdf"],
    },
    {
      id: uid("m"),
      name: "Lernplakat",
      beschreibung: "Ergebnissicherung und Präsentation in Gruppenarbeit auf Plakaten.",
      faecherIds: [],
      jahrgaenge: [8, 9],
      halbjahr: 1,
      materialien: [],
    },
  ];
}

// Beispielhafte Zuordnungen für eine Klasse, damit der Prototyp beim Start nicht leer ist
function seedPlanungen(klassen, lerngruppen, methoden) {
  const k8b = klassen.find((k) => k.jahrgang === 8 && k.buchstabe === "b");
  if (!k8b) return [];
  const findLg = (fachId) => lerngruppen.find((g) => g.fachId === fachId && g.klassenIds.includes(k8b.id));
  const findM = (name) => methoden.find((m) => m.name === name);
  const entries = [
    { m: "Kugellager", fachId: "f1", quartal: 1, status: "erledigt", datum: "2026-09-18", notiz: "Guter Einstieg, alle Gruppen aktiv beteiligt." },
    { m: "Gruppenpuzzle", fachId: "f3", quartal: 2, status: "erledigt", datum: "2026-11-05", notiz: "" },
    { m: "Mindmap erstellen", fachId: "f5", quartal: 1, status: "ausstehend", datum: null, notiz: "" },
    { m: "Lernplakat", fachId: "f2", quartal: 2, status: "ausgefallen", datum: null, notiz: "Entfällt wegen Praktikumswoche." },
    { m: "Textmarkierung", fachId: "f6", quartal: 3, status: "ausstehend", datum: null, notiz: "" },
    { m: "Concept Map", fachId: "f5", quartal: 4, status: "ausstehend", datum: null, notiz: "" },
  ];
  return entries
    .map((e) => {
      const methode = findM(e.m);
      const gruppe = findLg(e.fachId);
      if (!methode || !gruppe) return null;
      return {
        id: uid("p"),
        methodeId: methode.id,
        lerngruppeId: gruppe.id,
        klasseId: k8b.id,
        quartal: e.quartal,
        status: e.status,
        datum: e.datum,
        notiz: e.notiz,
      };
    })
    .filter(Boolean);
}

// ---------- Small UI building blocks ----------
function Badge({ children, tone = "muted" }) {
  const map = {
    muted: { bg: T.paperAlt, fg: T.muted, bd: T.line },
    accent: { bg: T.accentSoft, fg: "#7A5518", bd: T.accent },
    success: { bg: T.successSoft, fg: T.success, bd: T.success },
    danger: { bg: T.dangerSoft, fg: T.danger, bd: T.danger },
    cancelled: { bg: T.greySoft, fg: "#6C6A5F", bd: T.grey },
  };
  const c = map[tone];
  return (
    <span
      className="mc-mono text-xs px-2 py-0.5 rounded border uppercase tracking-wide"
      style={{ background: c.bg, color: c.fg, borderColor: c.bd }}
    >
      {children}
    </span>
  );
}

function Button({ children, onClick, tone = "default", disabled, small }) {
  const styles = {
    default: { background: T.ink, color: T.paper },
    accent: { background: T.accent, color: "#2A1D06" },
    success: { background: T.success, color: T.paper },
    danger: { background: T.danger, color: T.paper },
    ghost: { background: "transparent", color: T.ink, border: `1px solid ${T.line}` },
  };
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`mc-body font-medium rounded transition disabled:opacity-40 disabled:cursor-not-allowed hover:opacity-85 ${
        small ? "text-xs px-2 py-1" : "text-sm px-3 py-1.5"
      }`}
      style={styles[tone]}
    >
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="mc-body text-xs font-medium" style={{ color: T.muted }}>
        {label}
      </span>
      {children}
    </label>
  );
}

function inputStyle() {
  return { borderColor: T.line, background: "white" };
}

function StatusDot({ status }) {
  const color = status === "erledigt" ? T.success : status === "ausgefallen" ? T.grey : T.accent;
  return <span className="inline-block rounded-full" style={{ width: 9, height: 9, background: color }} />;
}

// ---------- Größenveränderbares Bild für den Tiptap-Editor ----------
// Normale Tiptap-Bilder kennen keine Breite/Höhe. Diese Erweiterung ergänzt beides als
// Attribute (werden auch mit exportiert/gespeichert) und rendert das Bild über eine eigene
// Komponente, die per ResizeObserver erkennt, wenn jemand am nativen Ziehpunkt unten rechts
// die Größe ändert, und das Ergebnis zurück ins Dokumentmodell schreibt – sonst würde die neue
// Größe zwar angezeigt, aber beim Speichern/Exportieren wieder verloren gehen.
function BildNodeAnsicht({ node, updateAttributes }) {
  const bildRef = useRef(null);

  useEffect(() => {
    const el = bildRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const beobachter = new ResizeObserver((eintraege) => {
      for (const eintrag of eintraege) {
        const breite = Math.round(eintrag.contentRect.width);
        const hoehe = Math.round(eintrag.contentRect.height);
        if (breite > 0 && hoehe > 0 && (breite !== node.attrs.width || hoehe !== node.attrs.height)) {
          updateAttributes({ width: breite, height: hoehe });
        }
      }
    });
    beobachter.observe(el);
    return () => beobachter.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <NodeViewWrapper as="span" style={{ display: "inline-block", verticalAlign: "bottom" }}>
      <img
        ref={bildRef}
        src={node.attrs.src}
        alt={node.attrs.alt || ""}
        style={{
          width: node.attrs.width ? `${node.attrs.width}px` : "auto",
          height: node.attrs.height ? `${node.attrs.height}px` : "auto",
          resize: "both",
          overflow: "hidden",
          cursor: "move",
          display: "block",
          maxWidth: "100%",
          borderRadius: 6,
        }}
        draggable
      />
    </NodeViewWrapper>
  );
}

const GroessenveraenderbaresBild = TiptapImage.extend({
  addAttributes() {
    return {
      ...this.parent(),
      width: { default: null, renderHTML: (attrs) => (attrs.width ? { width: attrs.width } : {}) },
      height: { default: null, renderHTML: (attrs) => (attrs.height ? { height: attrs.height } : {}) },
    };
  },
  addNodeView() {
    return ReactNodeViewRenderer(BildNodeAnsicht);
  },
});

// ---------- Word-Export der Methodenbeschreibung (HTML -> .docx) ----------
function farbeZuHex(css) {
  if (!css) return undefined;
  const m = css.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return undefined;
  return [1, 2, 3].map((i) => Number(m[i]).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function slug(text) {
  return (
    (text || "methode")
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || "methode"
  );
}

// ---------- Methoden-Export/Import (JSON) ----------
// Fächer werden nicht über ihre interne ID referenziert (die ist nur innerhalb einer
// Sitzung stabil und ändert sich z.B. bei jedem Neuladen durch den Untis-Auto-Import neu),
// sondern über das Fach-Kürzel – das bleibt über Sitzungen hinweg vergleichbar.
function methodeZuExportobjekt(m, faecher) {
  return {
    name: m.name,
    beschreibung: m.beschreibung || "",
    jahrgaenge: m.jahrgaenge || [],
    faecherKuerzel: (m.faecherIds || []).map((fid) => faecher.find((f) => f.id === fid)?.kuerzel).filter(Boolean),
    halbjahr: m.halbjahr,
    materialien: m.materialien || [],
    links: m.links || [],
  };
}

// Baut aus einem oder mehreren Export-Objekten frische Methoden auf; fehlende Fächer
// (Kürzel, das es in diesem Browser/dieser Sitzung noch nicht gibt) werden dabei neu
// angelegt, vorhandene wiederverwendet - im ganzen Stapel konsistent, nicht pro Methode
// einzeln, damit dasselbe Kürzel nicht mehrfach neu angelegt wird.
function bauMethodenAusExport(objekte, faecherAusgangswert) {
  const faecherPool = [...faecherAusgangswert];
  const findeOderErstelleFach = (kuerzel) => {
    let f = faecherPool.find((x) => x.kuerzel === kuerzel);
    if (!f) {
      f = { id: uid("f"), name: kuerzel, kuerzel, quelle: "import" };
      faecherPool.push(f);
    }
    return f;
  };
  const neueMethoden = objekte.map((obj) => ({
    id: uid("m"),
    name: obj.name || "Importierte Methode",
    beschreibung: obj.beschreibung || "",
    jahrgaenge: obj.jahrgaenge || [],
    faecherIds: (obj.faecherKuerzel || []).map((k) => findeOderErstelleFach(k).id),
    halbjahr: obj.halbjahr || 1,
    materialien: obj.materialien || [],
    links: obj.links || [],
  }));
  return { faecherPool, neueMethoden };
}

function jsonDateiHerunterladen(objekt, dateiname) {
  const blob = new Blob([JSON.stringify(objekt, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = dateiname;
  a.click();
  URL.revokeObjectURL(url);
}

function dataUrlZuBytes(dataUrl) {
  const komma = dataUrl.indexOf(",");
  const typ = dataUrl.slice(5, komma).split(";")[0].split("/")[1] || "png";
  const bin = atob(dataUrl.slice(komma + 1));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { bytes, typ: typ === "jpeg" ? "jpg" : typ };
}

function beschreibungZuDocxAbsaetzen(html) {
  const dok = new DOMParser().parseFromString(html && html.trim() ? html : "<p></p>", "text/html");
  const absaetze = [];

  const inlineKinder = (node, stil) => {
    let teile = [];
    node.childNodes.forEach((kind) => {
      if (kind.nodeType === Node.TEXT_NODE) {
        if (kind.textContent) teile.push(new TextRun({ text: kind.textContent, ...stil }));
      } else if (kind.nodeType === Node.ELEMENT_NODE) {
        const tag = kind.tagName.toLowerCase();
        if (tag === "img") {
          const src = kind.getAttribute("src") || "";
          if (src.startsWith("data:image/")) {
            const { bytes, typ } = dataUrlZuBytes(src);
            const breite = parseInt(kind.style.width) || parseInt(kind.getAttribute("width")) || 320;
            const hoehe = parseInt(kind.style.height) || parseInt(kind.getAttribute("height")) || Math.round(breite * 0.75);
            teile.push(new ImageRun({ data: bytes, type: typ, transformation: { width: breite, height: hoehe } }));
          }
          return;
        }
        const neu = { ...stil };
        if (tag === "strong" || tag === "b") neu.bold = true;
        if (tag === "em" || tag === "i") neu.italics = true;
        if (tag === "u") neu.underline = {};
        if (tag === "s" || tag === "strike") neu.strike = true;
        const farbe = farbeZuHex(kind.style && kind.style.color);
        if (farbe) neu.color = farbe;
        teile = teile.concat(inlineKinder(kind, neu));
      }
    });
    return teile;
  };

  const bloecke = dok.body.children.length ? Array.from(dok.body.children) : [];
  bloecke.forEach((block) => {
    const tag = block.tagName.toLowerCase();
    if (tag === "h1" || tag === "h2") {
      absaetze.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: inlineKinder(block, {}) }));
    } else if (tag === "h3" || tag === "h4") {
      absaetze.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: inlineKinder(block, {}) }));
    } else if (tag === "ul" || tag === "ol") {
      Array.from(block.children).forEach((li, i) => {
        const teile = inlineKinder(li, {});
        if (tag === "ol") teile.unshift(new TextRun(`${i + 1}. `));
        absaetze.push(new Paragraph({ children: teile.length ? teile : [new TextRun("")], bullet: tag === "ul" ? { level: 0 } : undefined }));
      });
    } else {
      const teile = inlineKinder(block, {});
      absaetze.push(new Paragraph({ children: teile.length ? teile : [new TextRun("")] }));
    }
  });

  return absaetze.length ? absaetze : [new Paragraph({ children: [new TextRun("")] })];
}

// Rich-Text-Feld für die vollständige Methodenbeschreibung: reine Anzeige.
// Bearbeitet wird ausschließlich über BeschreibungBearbeitenModal (Zugriff per Bearbeiten-Symbol).
function Beschreibungsfeld({ value }) {
  return value ? (
    <div className="text-sm rt-content" style={{ color: T.ink }} dangerouslySetInnerHTML={{ __html: value }} />
  ) : (
    <p className="text-sm italic" style={{ color: T.muted }}>
      Keine Beschreibung hinterlegt.
    </p>
  );
}

// Großes Bearbeitungs-Fenster für die vollständige Methodenbeschreibung: Überschriften, Listen,
// Textfarbe, Bilder. Nutzt document.execCommand (bleibt dadurch auch in dieser
// Vorschau nutzbar) und wird über ein Bearbeiten-Symbol in der Methoden-Übersicht geöffnet.
function BeschreibungBearbeitenModal({ title, value, onChange, onClose }) {
  const bildInputRef = useRef(null);
  const wordImportRef = useRef(null);
  const [exportLaeuft, setExportLaeuft] = useState(false);

  const editor = useEditor(
    {
      extensions: [
        StarterKit.configure({ heading: { levels: [2, 3] } }),
        GroessenveraenderbaresBild,
        TextStyle,
        Color,
        Placeholder.configure({ placeholder: "Beschreibung der Methode…" }),
      ],
      content: value || "",
      onUpdate: ({ editor }) => onChange(editor.getHTML()),
      editorProps: {
        attributes: {
          class: "rt-content text-sm rounded border px-4 py-4 focus:outline-none",
        },
      },
    },
    // Leeres deps-Array: Der Editor wird nur einmal beim Öffnen mit "value" erzeugt und danach
    // nicht erneut synchronisiert – sonst würde jede Änderung (die ja über onChange den Zustand
    // weiter oben aktualisiert) den Editor zurücksetzen und den Cursor springen lassen.
    []
  );

  const toolbarBtn = (label, title, onClick, aktiv, style) => (
    <button
      type="button"
      title={title}
      onMouseDown={(e) => {
        e.preventDefault();
        onClick();
      }}
      className="text-xs min-w-[26px] h-7 px-1.5 rounded border flex items-center justify-center"
      style={{ borderColor: aktiv ? T.accent : T.line, background: aktiv ? T.accentSoft : "white", color: T.ink, ...style }}
    >
      {label}
    </button>
  );

  // Bild wird zunächst geladen, um die natürliche Größe zu kennen (auf max. Breite begrenzt) –
  // erst mit fester Breite/Höhe lässt es sich später per Eckziehen verändern (siehe
  // GroessenveraenderbaresBild) und beim Word-Export korrekt einbetten.
  const bildEinfuegen = (e) => {
    const datei = e.target.files[0];
    if (!datei || !editor) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const bild = new window.Image();
      bild.onload = () => {
        const maxBreite = 420;
        const breite = Math.min(bild.naturalWidth || maxBreite, maxBreite);
        const hoehe = bild.naturalWidth ? Math.round((bild.naturalHeight / bild.naturalWidth) * breite) : Math.round(breite * 0.75);
        editor.chain().focus().setImage({ src: dataUrl, width: breite, height: hoehe }).run();
      };
      bild.src = dataUrl;
    };
    reader.readAsDataURL(datei);
    e.target.value = "";
  };

  const wordImportieren = (e) => {
    const datei = e.target.files[0];
    if (!datei || !editor) return;
    const reader = new FileReader();
    reader.onload = () => {
      mammoth
        .convertToHtml(
          { arrayBuffer: reader.result },
          { convertImage: mammoth.images.imgElement((bild) => bild.read("base64").then((daten) => ({ src: "data:" + bild.contentType + ";base64," + daten }))) }
        )
        .then((ergebnis) => {
          editor.commands.setContent(ergebnis.value);
          onChange(ergebnis.value);
        })
        .catch((err) => console.error("Word-Import fehlgeschlagen:", err));
    };
    reader.readAsArrayBuffer(datei);
    e.target.value = "";
  };

  const alsWordExportieren = async () => {
    setExportLaeuft(true);
    try {
      const absaetze = beschreibungZuDocxAbsaetzen(value);
      const doc = new Document({ sections: [{ children: absaetze }] });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${slug(title)}.docx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Word-Export fehlgeschlagen:", err);
    } finally {
      setExportLaeuft(false);
    }
  };

  return (
    <div className="fixed inset-y-0 right-0 left-0 lg:left-64 z-50 flex flex-col" style={{ background: T.paper }}>
      <div className="flex items-center justify-between gap-3 px-6 py-3 border-b" style={{ borderColor: T.line, background: "white" }}>
        <h2 className="mc-display text-lg font-semibold truncate">{title}</h2>
        <div className="flex items-center gap-2 shrink-0">
          <Button small tone="ghost" onClick={() => wordImportRef.current?.click()}>
            Word importieren
          </Button>
          <input ref={wordImportRef} type="file" accept=".docx" onChange={wordImportieren} className="hidden" />
          <Button small tone="ghost" onClick={alsWordExportieren} disabled={exportLaeuft}>
            {exportLaeuft ? "Exportiere…" : "Als Word exportieren"}
          </Button>
          <Button small onClick={onClose}>
            Fertig
          </Button>
        </div>
      </div>

      {editor && (
        <div className="flex items-center gap-1 px-6 py-2 border-b flex-wrap" style={{ borderColor: T.line, background: "white" }}>
          {toolbarBtn("F", "Fett", () => editor.chain().focus().toggleBold().run(), editor.isActive("bold"), { fontWeight: 700 })}
          {toolbarBtn("K", "Kursiv", () => editor.chain().focus().toggleItalic().run(), editor.isActive("italic"), { fontStyle: "italic" })}
          {toolbarBtn("U", "Unterstrichen", () => editor.chain().focus().toggleUnderline().run(), editor.isActive("underline"), { textDecoration: "underline" })}
          {toolbarBtn("D", "Durchgestrichen", () => editor.chain().focus().toggleStrike().run(), editor.isActive("strike"), { textDecoration: "line-through" })}
          <span className="mx-1" style={{ width: 1, height: 20, background: T.line }} />
          {toolbarBtn("H2", "Überschrift groß", () => editor.chain().focus().toggleHeading({ level: 2 }).run(), editor.isActive("heading", { level: 2 }), { fontWeight: 700 })}
          {toolbarBtn("H3", "Überschrift klein", () => editor.chain().focus().toggleHeading({ level: 3 }).run(), editor.isActive("heading", { level: 3 }), { fontWeight: 700 })}
          {toolbarBtn("T", "Fließtext", () => editor.chain().focus().setParagraph().run(), editor.isActive("paragraph"), {})}
          <span className="mx-1" style={{ width: 1, height: 20, background: T.line }} />
          {toolbarBtn("≡", "Aufzählung", () => editor.chain().focus().toggleBulletList().run(), editor.isActive("bulletList"), {})}
          {toolbarBtn("1.", "Nummerierung", () => editor.chain().focus().toggleOrderedList().run(), editor.isActive("orderedList"), {})}
          <span className="mx-1" style={{ width: 1, height: 20, background: T.line }} />
          {[T.ink, T.accent, T.success, T.danger].map((c) => (
            <button
              key={c}
              type="button"
              title="Textfarbe"
              onMouseDown={(e) => {
                e.preventDefault();
                editor.chain().focus().setColor(c).run();
              }}
              className="w-5 h-5 rounded-full border"
              style={{ background: c, borderColor: T.line }}
            />
          ))}
          <span className="mx-1" style={{ width: 1, height: 20, background: T.line }} />
          <button
            type="button"
            title="Bild einfügen (danach per Eckziehen im Text größenveränderbar, per Ziehen verschiebbar)"
            onMouseDown={(e) => {
              e.preventDefault();
              bildInputRef.current?.click();
            }}
            className="text-xs h-7 px-2 rounded border"
            style={{ borderColor: T.line, color: T.ink }}
          >
            🖼 Bild
          </button>
          <input ref={bildInputRef} type="file" accept="image/*" onChange={bildEinfuegen} className="hidden" />
        </div>
      )}

      <div className="flex-1 overflow-y-auto px-6 py-5">
        <EditorContent editor={editor} />
      </div>
    </div>
  );
}

function MethodCard({ planung, methode, onClick, onDragStart }) {
  const toneByStatus = {
    ausstehend: { border: T.accent, bg: "#FFFCF6", text: T.ink, strike: false },
    erledigt: { border: T.success, bg: T.successSoft, text: T.ink, strike: false },
    ausgefallen: { border: T.grey, bg: T.greySoft, text: T.muted, strike: true },
  };
  const s = toneByStatus[planung.status];
  // Erledigte Methoden sind dokumentiert (Datum, Notiz) und sollen nicht versehentlich
  // verschoben werden können – erst nach einem Statuswechsel wieder ziehbar.
  const kannGezogenWerden = !!onDragStart && planung.status !== "erledigt";
  return (
    <button
      onClick={onClick}
      draggable={kannGezogenWerden}
      onDragStart={kannGezogenWerden ? onDragStart : undefined}
      title={kannGezogenWerden ? methode.name : `${methode.name} – Status ändern, um zu verschieben`}
      className={`text-left rounded px-2.5 py-1.5 text-xs font-medium transition hover:-translate-y-0.5 ${
        kannGezogenWerden ? "cursor-grab active:cursor-grabbing" : ""
      }`}
      style={{
        border: `1.5px solid ${s.border}`,
        background: s.bg,
        color: s.text,
        textDecoration: s.strike ? "line-through" : "none",
        maxWidth: "100%",
      }}
    >
      {methode.name}
    </button>
  );
}

// ---------- App ----------
export default function App() {
  const [ladezustand, setLadezustand] = useState("laedt"); // 'laedt' | 'bereit' | 'fehler'
  const [ladeFehler, setLadeFehler] = useState("");
  const [datenquelle, setDatenquelle] = useState(null); // 'supabase' | 'lokal' (noch nicht nach Supabase übernommen)

  const [faecher, setFaecher] = useState([]);
  const [lehrer, setLehrer] = useState([]);
  const [klassen, setKlassen] = useState([]);
  const [lerngruppen, setLerngruppen] = useState([]);
  const [methoden, setMethoden] = useState([]);
  const [planungen, setPlanungen] = useState([]);

  useEffect(() => {
    let abgebrochen = false;
    db.ladeAlleDaten()
      .then((daten) => {
        if (abgebrochen) return;
        if (daten.klassen.length === 0) {
          // Noch keine Migration durchgeführt (leere Supabase-Tabellen) - mit dem bisherigen
          // lokalen Ausgangszustand starten, bis "In Supabase übernehmen" ausgeführt wurde.
          const lokal = seedInitialerZustand();
          setFaecher(lokal.faecher);
          setLehrer(lokal.lehrer);
          setKlassen(lokal.klassen);
          setLerngruppen(lokal.lerngruppen);
          setMethoden(seedMethoden());
          setPlanungen([]);
          setDatenquelle("lokal");
          setSelectedKlasseId((lokal.klassen.find((k) => k.jahrgang === 8 && k.buchstabe === "b") || {}).id || null);
        } else {
          setFaecher(daten.faecher);
          setLehrer(daten.lehrer);
          setKlassen(daten.klassen);
          setLerngruppen(daten.lerngruppen);
          setMethoden(daten.methoden);
          setPlanungen(daten.planungen);
          setDatenquelle("supabase");
          setSelectedKlasseId((daten.klassen.find((k) => k.jahrgang === 8 && k.buchstabe === "b") || daten.klassen[0] || {}).id || null);
        }
        setLadezustand("bereit");
      })
      .catch((err) => {
        if (abgebrochen) return;
        console.error("Laden aus Supabase fehlgeschlagen:", err);
        setLadeFehler(err.message || String(err));
        setLadezustand("fehler");
      });
    return () => {
      abgebrochen = true;
    };
  }, []);

  const [mode, setMode] = useState("klasse"); // 'klasse' | 'lehrer' | 'verwaltung'
  const [selectedKlasseId, setSelectedKlasseId] = useState(null);
  const [expanded, setExpanded] = useState({ 8: true });
  const [stammTab, setStammTab] = useState("faecher");
  const [sidebarOffen, setSidebarOffen] = useState(false); // nur auf schmalen Ansichten (iPad Hochformat) relevant
  const [currentLehrerId, setCurrentLehrerId] = useState(null);
  const [modalPlanungId, setModalPlanungId] = useState(null);

  useEffect(() => {
    if (lehrer.length && !currentLehrerId) setCurrentLehrerId(lehrer[0].id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lehrer]);

  // ----- lookups -----
  const fach = (id) => faecher.find((f) => f.id === id);
  const lehr = (id) => lehrer.find((l) => l.id === id);
  const lg = (id) => lerngruppen.find((g) => g.id === id);
  const meth = (id) => methoden.find((m) => m.id === id);
  const klass = (id) => klassen.find((k) => k.id === id);

  const toggleJahrgang = (jg) => setExpanded((s) => ({ ...s, [jg]: !s[jg] }));
  const waehleKlasse = (id) => {
    setSelectedKlasseId(id);
    setMode("klasse");
    setSidebarOffen(false); // auf schmalen Ansichten nach Auswahl automatisch schließen
  };
  const waehleModus = (m) => {
    setMode(m);
    setSidebarOffen(false);
  };

  // ----- Stammdaten: Fächer -----
  const [neuFach, setNeuFach] = useState({ name: "", kuerzel: "" });
  const addFach = () => {
    if (!neuFach.name || !neuFach.kuerzel) return;
    const eingabe = neuFach;
    const neu = { id: uid("f"), ...eingabe };
    setFaecher([...faecher, neu]);
    setNeuFach({ name: "", kuerzel: "" });
    if (datenquelle === "supabase") {
      db.fachErstellen(eingabe)
        .then((serverFach) => setFaecher((prev) => prev.map((f) => (f.id === neu.id ? serverFach : f))))
        .catch((err) => console.error("Fach anlegen fehlgeschlagen:", err));
    }
  };

  // ----- Stammdaten: Lehrer -----
  const [neuLehrer, setNeuLehrer] = useState({ name: "", kuerzel: "", email: "" });
  const addLehrer = () => {
    if (!neuLehrer.name || neuLehrer.kuerzel.length !== 3) return;
    const eingabe = neuLehrer;
    const neu = { id: uid("l"), ...eingabe };
    setLehrer([...lehrer, neu]);
    setNeuLehrer({ name: "", kuerzel: "", email: "" });
    if (datenquelle === "supabase") {
      db.lehrerErstellen(eingabe)
        .then((serverLehrer) => setLehrer((prev) => prev.map((l) => (l.id === neu.id ? serverLehrer : l))))
        .catch((err) => console.error("Lehrkraft anlegen fehlgeschlagen:", err));
    }
  };

  // ----- Stammdaten: Klassen -----
  const updateKlasse = (id, patch) => {
    setKlassen(klassen.map((k) => (k.id === id ? { ...k, ...patch } : k)));
    if (datenquelle === "supabase") {
      db.klasseAktualisieren(id, patch).catch((err) => console.error("Klasse aktualisieren fehlgeschlagen:", err));
    }
  };
  const [neuKlasse, setNeuKlasse] = useState({ jahrgang: 5, buchstabe: "a" });
  const addKlasse = () => {
    const eingabe = { jahrgang: Number(neuKlasse.jahrgang), buchstabe: neuKlasse.buchstabe, lehrer1: null, lehrer2: null };
    const neu = { id: uid("k"), ...eingabe, vorgaenger: null };
    setKlassen([...klassen, neu]);
    if (datenquelle === "supabase") {
      db.klasseErstellen(eingabe)
        .then((serverKlasse) => setKlassen((prev) => prev.map((k) => (k.id === neu.id ? serverKlasse : k))))
        .catch((err) => console.error("Klasse anlegen fehlgeschlagen:", err));
    }
  };

  // ----- Stammdaten: Lerngruppen -----
  const [neuLg, setNeuLg] = useState({ fachId: "f1", bezeichnung: "", jahrgang: 5, lehrerId: "", klassenIds: [] });
  const addLerngruppe = () => {
    if (!neuLg.bezeichnung || !neuLg.lehrerId) return;
    const eingabe = { ...neuLg, jahrgang: Number(neuLg.jahrgang) };
    const neu = { id: uid("lg"), ...eingabe };
    setLerngruppen([...lerngruppen, neu]);
    setNeuLg({ fachId: "f1", bezeichnung: "", jahrgang: 5, lehrerId: "", klassenIds: [] });
    if (datenquelle === "supabase") {
      db.lerngruppeErstellen(eingabe)
        .then((serverGruppe) => setLerngruppen((prev) => prev.map((g) => (g.id === neu.id ? serverGruppe : g))))
        .catch((err) => console.error("Lerngruppe anlegen fehlgeschlagen:", err));
    }
  };
  const toggleKlasseInNeuLg = (kid) => {
    setNeuLg((s) => ({
      ...s,
      klassenIds: s.klassenIds.includes(kid) ? s.klassenIds.filter((x) => x !== kid) : [...s.klassenIds, kid],
    }));
  };
  const [lerngruppenSicherung, setLerngruppenSicherung] = useState(null); // { zustand, beschreibung }
  const updateLerngruppe = (id, patch) => {
    const g = lerngruppen.find((x) => x.id === id);
    setLerngruppenSicherung({ zustand: lerngruppen, beschreibung: `Lehrerwechsel bei "${g?.bezeichnung || ""}"` });
    setLerngruppen((prev) => prev.map((x) => (x.id === id ? { ...x, ...patch } : x)));
    if (datenquelle === "supabase") {
      db.lerngruppeAktualisieren(id, patch).catch((err) => console.error("Lerngruppe aktualisieren fehlgeschlagen:", err));
    }
  };
  const removeLerngruppe = (id) => {
    const g = lerngruppen.find((x) => x.id === id);
    setLerngruppenSicherung({ zustand: lerngruppen, beschreibung: `Löschen von "${g?.bezeichnung || ""}"` });
    setLerngruppen((prev) => prev.filter((x) => x.id !== id));
    if (datenquelle === "supabase") {
      db.lerngruppeLoeschen(id).catch((err) => console.error("Lerngruppe löschen fehlgeschlagen:", err));
    }
  };
  const undoLerngruppenAenderung = () => {
    if (!lerngruppenSicherung) return;
    setLerngruppen(lerngruppenSicherung.zustand);
    setLerngruppenSicherung(null);
    // Absichtlich kein Supabase-Rückgängig: "Rückgängig" wirkt bewusst nur lokal/sofort;
    // die vorherige Datenbank-Änderung müsste sonst erneut nachvollzogen werden.
  };
  const verwirfLerngruppenSicherung = () => setLerngruppenSicherung(null);

  // ----- Methoden -----
  const [neuMethode, setNeuMethode] = useState({ name: "", beschreibung: "", faecherIds: [], jahrgaenge: [], halbjahr: 1 });
  const [neuMethodeKey, setNeuMethodeKey] = useState(0);
  const addMethode = () => {
    if (!neuMethode.name || neuMethode.faecherIds.length === 0 || neuMethode.jahrgaenge.length === 0) return;
    const eingabe = { materialien: [], links: [], ...neuMethode };
    const neu = { id: uid("m"), ...eingabe };
    setMethoden([...methoden, neu]);
    setNeuMethode({ name: "", beschreibung: "", faecherIds: [], jahrgaenge: [], halbjahr: 1 });
    setNeuMethodeKey((k) => k + 1); // erzwingt Reset des unkontrollierten Rich-Text-Felds
    if (datenquelle === "supabase") {
      db.methodeErstellen(eingabe)
        .then((serverMethode) => setMethoden((prev) => prev.map((m) => (m.id === neu.id ? serverMethode : m))))
        .catch((err) => console.error("Methode anlegen fehlgeschlagen:", err));
    }
  };
  const toggleInList = (list, val) => (list.includes(val) ? list.filter((x) => x !== val) : [...list, val]);
  const updateMethode = (id, patch) => {
    setMethoden((prev) => prev.map((m) => (m.id === id ? { ...m, ...patch } : m)));
    if (datenquelle === "supabase") {
      db.methodeAktualisieren(id, patch).catch((err) => console.error("Methode aktualisieren fehlgeschlagen:", err));
    }
  };
  // Löschen einer Methode kaskadiert auf ihre Planungen (Zeitleisten-Zuordnungen) - sonst
  // blieben dort tote Verweise auf eine nicht mehr existierende Methode zurück. Bei einer
  // echten Datenbank übernimmt das "on delete cascade" im Schema automatisch mit.
  const removeMethode = (id) => {
    setMethoden((prev) => prev.filter((m) => m.id !== id));
    setPlanungen((prev) => prev.filter((p) => p.methodeId !== id));
    if (datenquelle === "supabase") {
      db.methodeLoeschen(id).catch((err) => console.error("Methode löschen fehlgeschlagen:", err));
    }
  };

  // ----- Planung (Zuordnung + Durchführung in einem) -----
  const lerngruppenFuerKlasse = (klasseId) => lerngruppen.filter((g) => g.klassenIds.includes(klasseId));
  const planungenFuerKlasse = (klasseId) => planungen.filter((p) => p.klasseId === klasseId);
  const planungenFuerLehrer = (lehrerId) =>
    planungen.filter((p) => {
      const g = lg(p.lerngruppeId);
      return g && g.lehrerId === lehrerId;
    });

  const vorschlaegeFuerKlasse = (klasseId) => {
    const k = klass(klasseId);
    if (!k) return [];
    const bereitsPlatziert = new Set(planungenFuerKlasse(klasseId).map((p) => p.methodeId));
    return methoden
      .filter((m) => m.jahrgaenge.includes(k.jahrgang) && !bereitsPlatziert.has(m.id))
      .map((m) => {
        const passendeGruppe = m.faecherIds
          .map((fid) => lerngruppen.find((g) => g.fachId === fid && g.klassenIds.includes(klasseId)))
          .find(Boolean);
        return passendeGruppe ? { methode: m, lerngruppe: passendeGruppe } : null;
      })
      .filter(Boolean);
  };

  const platziereMethode = (klasseId, methodeId, lerngruppeId, quartal) => {
    const eingabe = { methodeId, lerngruppeId, klasseId, quartal, status: "ausstehend", datum: null, notiz: "" };
    const neu = { id: uid("p"), ...eingabe };
    setPlanungen((prev) => [...prev, neu]);
    if (datenquelle === "supabase") {
      db.planungErstellen(eingabe)
        .then((serverPlanung) => setPlanungen((prev) => prev.map((p) => (p.id === neu.id ? serverPlanung : p))))
        .catch((err) => console.error("Planung anlegen fehlgeschlagen:", err));
    }
  };

  const updatePlanung = (id, patch) => {
    setPlanungen((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
    if (datenquelle === "supabase") {
      db.planungAktualisieren(id, patch).catch((err) => console.error("Planung aktualisieren fehlgeschlagen:", err));
    }
  };
  const removePlanung = (id) => {
    setPlanungen((prev) => prev.filter((p) => p.id !== id));
    setModalPlanungId(null);
    if (datenquelle === "supabase") {
      db.planungLoeschen(id).catch((err) => console.error("Planung löschen fehlgeschlagen:", err));
    }
  };

  const setzeStatus = (id, status) => {
    let neuesDatum;
    setPlanungen((prev) =>
      prev.map((p) => {
        if (p.id !== id) return p;
        const datum = status === "erledigt" ? p.datum || new Date().toISOString().slice(0, 10) : p.datum;
        neuesDatum = datum;
        return { ...p, status, datum };
      })
    );
    if (datenquelle === "supabase") {
      db.planungAktualisieren(id, { status, datum: neuesDatum }).catch((err) => console.error("Status aktualisieren fehlgeschlagen:", err));
    }
  };

  // ----- Drag & Drop -----
  // Ein "chip" kommt aus dem Empfehlungen-Feld (noch nicht zugeordnet), eine "card" ist
  // bereits platziert und wird von einem Fach/Quartal in ein anderes gezogen.
  const onDragStartChip = (e, methodeId, lerngruppeId) => {
    e.dataTransfer.setData("text/plain", JSON.stringify({ typ: "chip", methodeId, lerngruppeId }));
  };
  const onDragStartCard = (e, planungId) => {
    e.dataTransfer.setData("text/plain", JSON.stringify({ typ: "card", planungId }));
  };
  const onDropLane = (e, klasseId, fachId, quartal) => {
    e.preventDefault();
    let data;
    try {
      data = JSON.parse(e.dataTransfer.getData("text/plain"));
    } catch {
      return;
    }
    if (!data) return;
    const gruppe = lerngruppen.find((g) => g.fachId === fachId && g.klassenIds.includes(klasseId));
    if (!gruppe) return;
    if (data.typ === "card" && data.planungId) {
      updatePlanung(data.planungId, { lerngruppeId: gruppe.id, quartal });
      return;
    }
    if (data.typ === "chip" && data.methodeId) {
      platziereMethode(klasseId, data.methodeId, gruppe.id, quartal);
    }
  };
  const onDropZurueckZuEmpfehlungen = (e) => {
    e.preventDefault();
    let data;
    try {
      data = JSON.parse(e.dataTransfer.getData("text/plain"));
    } catch {
      return;
    }
    if (data && data.typ === "card" && data.planungId) {
      removePlanung(data.planungId);
    }
  };

  // ----- Manuelle Zuordnung (Fallback ohne Drag & Drop) -----
  const [manuell, setManuell] = useState({ methodeId: "", lerngruppeId: "", quartal: 1 });
  const addManuell = (klasseId) => {
    if (!manuell.methodeId || !manuell.lerngruppeId) return;
    platziereMethode(klasseId, manuell.methodeId, manuell.lerngruppeId, Number(manuell.quartal));
    setManuell({ methodeId: "", lerngruppeId: "", quartal: 1 });
  };

  // ----- Einmalige Migration nach Supabase -----
  const [migrationLaeuft, setMigrationLaeuft] = useState(false);
  const [migrationFortschritt, setMigrationFortschritt] = useState("");
  const [migrationFehler, setMigrationFehler] = useState("");
  const migriereNachSupabase = async () => {
    setMigrationLaeuft(true);
    setMigrationFehler("");
    try {
      const leer = await db.supabaseIstLeer();
      if (!leer) {
        setMigrationFehler(
          "In Supabase liegen bereits Daten (z.B. Rest eines vorherigen, abgebrochenen Versuchs). " +
            "Bitte erst die Tabellen leeren (siehe supabase-reset.sql) und dann erneut versuchen."
        );
        return;
      }
      await db.allesNachSupabaseHochladen({ faecher, lehrer, klassen, lerngruppen, methoden, planungen }, setMigrationFortschritt);
      const daten = await db.ladeAlleDaten();
      setFaecher(daten.faecher);
      setLehrer(daten.lehrer);
      setKlassen(daten.klassen);
      setLerngruppen(daten.lerngruppen);
      setMethoden(daten.methoden);
      setPlanungen(daten.planungen);
      setDatenquelle("supabase");
    } catch (err) {
      console.error("Migration nach Supabase fehlgeschlagen:", err);
      setMigrationFehler(err.message || String(err));
    } finally {
      setMigrationLaeuft(false);
      setMigrationFortschritt("");
    }
  };

  const modalPlanung = modalPlanungId ? planungen.find((p) => p.id === modalPlanungId) : null;

  // ---------- Render ----------
  if (ladezustand === "laedt") {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: T.paper, color: T.muted }}>
        <style>{FONTS}</style>
        <p className="mc-body text-sm">Daten werden geladen…</p>
      </div>
    );
  }
  if (ladezustand === "fehler") {
    return (
      <div className="min-h-screen flex items-center justify-center p-8" style={{ background: T.paper }}>
        <style>{FONTS}</style>
        <div className="max-w-md rounded-xl border p-6 bg-white" style={{ borderColor: T.danger }}>
          <div className="mc-display text-lg font-semibold mb-2" style={{ color: T.danger }}>
            Verbindung zu Supabase fehlgeschlagen
          </div>
          <p className="text-sm mb-1" style={{ color: T.ink }}>
            {ladeFehler}
          </p>
          <p className="text-xs" style={{ color: T.muted }}>
            Prüfe die Adresse und den Schlüssel in src/supabaseClient.js sowie die Zugriffsregeln (RLS) in Supabase.
          </p>
        </div>
      </div>
    );
  }
  return (
    <div className="mc-body min-h-screen flex" style={{ background: T.paper, color: T.ink }}>
      <style>{FONTS}</style>

      {/* Abdunklung hinter der Seitenleiste, nur auf schmalen Ansichten (iPad Hochformat) sichtbar */}
      {sidebarOffen && (
        <div
          className="fixed inset-0 z-30 lg:hidden"
          style={{ background: "rgba(0,0,0,0.4)" }}
          onClick={() => setSidebarOffen(false)}
        />
      )}

      {/* ---------------- SIDEBAR ---------------- */}
      {/* Ab "lg" (≈ Tablet-Querformat und größer) fest sichtbar wie bisher; darunter (z.B.
          iPad im Hochformat) ein einblendbares Overlay, das über den Hamburger-Knopf in
          "MAIN" geöffnet wird und nach einer Auswahl automatisch wieder zuklappt. */}
      <aside
        className={`fixed inset-y-0 left-0 z-40 w-72 lg:w-64 lg:static lg:translate-x-0 shrink-0 flex flex-col transition-transform duration-200 ${
          sidebarOffen ? "translate-x-0" : "-translate-x-full"
        }`}
        style={{ background: T.board, color: T.boardText }}
      >
        <div className="px-4 py-4 border-b" style={{ borderColor: T.boardLight }}>
          <div className="mc-mono text-[11px] uppercase tracking-wide" style={{ color: T.boardTextMuted }}>
            Schulplattform
          </div>
          <div className="mc-display text-lg font-semibold">Methodencurriculum</div>
          <div className="mc-mono text-xs mt-0.5" style={{ color: T.boardTextMuted }}>
            Schuljahr 2026/27 · Prototyp
          </div>
        </div>

        <div className="px-2 pt-2">
          <div
            className="mx-2 mb-1 px-3 py-1.5 rounded mc-body text-xs cursor-not-allowed flex items-center justify-between"
            style={{ color: T.boardTextMuted }}
            title="Bereits als eigenständige App vorhanden – Anbindung folgt"
          >
            Raumplanung <span>bald</span>
          </div>
          <div
            className="mx-2 mb-2 px-3 py-1.5 rounded mc-body text-xs cursor-not-allowed flex items-center justify-between"
            style={{ color: T.boardTextMuted }}
            title="Geplantes Modul"
          >
            Umfragen <span>bald</span>
          </div>
        </div>

        <nav className="px-2">
          <button
            onClick={() => waehleModus("lehrer")}
            className="w-full text-left mc-body text-sm px-3 py-2.5 rounded mb-1"
            style={mode === "lehrer" ? { background: T.boardActive, fontWeight: 600 } : { color: "#DCE3DE" }}
          >
            👤 Meine Methoden
          </button>
          <button
            onClick={() => waehleModus("verwaltung")}
            className="w-full text-left mc-body text-sm px-3 py-2.5 rounded mb-1"
            style={mode === "verwaltung" ? { background: T.boardActive, fontWeight: 600 } : { color: "#DCE3DE" }}
          >
            ⚙ Verwaltung
          </button>
        </nav>

        <div className="px-2 mt-1 flex-1 overflow-y-auto">
          <div className="mc-mono text-[11px] uppercase tracking-wide px-3 pt-2 pb-1" style={{ color: T.boardTextMuted }}>
            Klassen
          </div>
          {JAHRGAENGE.map((jg) => (
            <div key={jg} className="mb-0.5">
              <button
                onClick={() => toggleJahrgang(jg)}
                className="w-full flex items-center gap-2 text-left mc-display text-sm font-medium px-2 py-2.5 rounded"
                style={{ color: "#DCE3DE" }}
              >
                <span
                  className="inline-block text-[10px] transition-transform"
                  style={{ color: T.accent, transform: expanded[jg] ? "rotate(90deg)" : "rotate(0deg)" }}
                >
                  ▸
                </span>
                Jahrgang {jg}
              </button>
              {expanded[jg] && (
                <div className="pl-5">
                  {BUCHSTABEN.map((b) => {
                    const k = klassen.find((kk) => kk.jahrgang === jg && kk.buchstabe === b);
                    if (!k) return null;
                    const aktiv = mode === "klasse" && selectedKlasseId === k.id;
                    return (
                      <div
                        key={k.id}
                        onClick={() => waehleKlasse(k.id)}
                        className="flex items-center gap-2 text-sm px-2.5 py-2.5 rounded cursor-pointer"
                        style={aktiv ? { background: T.boardActive, color: "#fff", fontWeight: 600 } : { color: "#C7D0CB" }}
                      >
                        <span
                          className="inline-block rounded-full"
                          style={{ width: 5, height: 5, background: aktiv ? T.accent : "transparent" }}
                        />
                        {jg}
                        {b}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          ))}
        </div>

        <div className="px-4 py-3 border-t" style={{ borderColor: T.boardLight }}>
          <div className="mc-mono text-[10px] uppercase tracking-wide mb-1" style={{ color: T.boardTextMuted }}>
            Angemeldet als (Simulation)
          </div>
          <select
            className="w-full text-sm rounded px-2 py-1"
            style={{ background: T.boardLight, color: "#fff", border: "none" }}
            value={currentLehrerId}
            onChange={(e) => setCurrentLehrerId(e.target.value)}
          >
            {lehrer.length === 0 ? (
              <option value="">— keine Lehrkräfte angelegt —</option>
            ) : (
              lehrer.map((l) => (
                <option key={l.id} value={l.id}>
                  {l.name || l.kuerzel}
                </option>
              ))
            )}
          </select>
        </div>
      </aside>

      {/* ---------------- MAIN ---------------- */}
      <main className="flex-1 min-w-0 overflow-y-auto flex flex-col">
        {/* Kopfleiste nur auf schmalen Ansichten (iPad Hochformat) – öffnet die Seitenleiste */}
        <div
          className="lg:hidden flex items-center gap-3 px-4 py-3 border-b sticky top-0 z-20 shrink-0"
          style={{ background: T.paper, borderColor: T.line }}
        >
          <button
            onClick={() => setSidebarOffen(true)}
            className="w-11 h-11 flex items-center justify-center rounded-lg border shrink-0"
            style={{ borderColor: T.line, background: "white" }}
            aria-label="Menü öffnen"
          >
            <span className="text-xl leading-none">☰</span>
          </button>
          <span className="mc-display text-sm font-semibold truncate">Methodencurriculum</span>
        </div>

        <div className="p-4 md:p-8 flex-1">
        {mode === "verwaltung" && (
          <VerwaltungView
            {...{
              faecher, lehrer, klassen, lerngruppen, methoden,
              setFaecher, setLehrer, setKlassen, setLerngruppen, setMethoden,
              planungen, setPlanungen,
              stammTab, setStammTab,
              neuFach, setNeuFach, addFach,
              neuLehrer, setNeuLehrer, addLehrer,
              updateKlasse, neuKlasse, setNeuKlasse, addKlasse,
              neuLg, setNeuLg, addLerngruppe, toggleKlasseInNeuLg, updateLerngruppe, removeLerngruppe,
              lerngruppenSicherung, undoLerngruppenAenderung, verwirfLerngruppenSicherung,
              neuMethode, setNeuMethode, addMethode, toggleInList, updateMethode, removeMethode, neuMethodeKey,
              datenquelle, migriereNachSupabase, migrationLaeuft, migrationFortschritt, migrationFehler,
            }}
          />
        )}

        {mode === "lehrer" && (
          <LehrerAnsicht
            lehrerObj={lehr(currentLehrerId)}
            planungen={planungenFuerLehrer(currentLehrerId)}
            klass={klass}
            meth={meth}
            fach={fach}
            lg={lg}
            onOeffne={setModalPlanungId}
          />
        )}

        {mode === "klasse" && selectedKlasseId && (
          <KlasseAnsicht
            klasse={klass(selectedKlasseId)}
            lehr={lehr}
            fach={fach}
            meth={meth}
            lg={lg}
            lerngruppenFuerKlasse={lerngruppenFuerKlasse(selectedKlasseId)}
            planungenFuerKlasse={planungenFuerKlasse(selectedKlasseId)}
            vorschlaege={vorschlaegeFuerKlasse(selectedKlasseId)}
            onDragStartChip={onDragStartChip}
            onDragStartCard={onDragStartCard}
            onDropLane={onDropLane}
            onDropZurueckZuEmpfehlungen={onDropZurueckZuEmpfehlungen}
            onOeffne={setModalPlanungId}
            manuell={manuell}
            setManuell={setManuell}
            addManuell={addManuell}
            methoden={methoden}
            lerngruppen={lerngruppen}
          />
        )}

        {mode === "klasse" && !selectedKlasseId && (
          <p className="text-sm" style={{ color: T.muted }}>
            Bitte links eine Klasse auswählen.
          </p>
        )}
        </div>
      </main>

      {/* ---------------- DETAIL-MODAL ---------------- */}
      {modalPlanung && (
        <DetailModal
          planung={modalPlanung}
          methode={meth(modalPlanung.methodeId)}
          gruppe={lg(modalPlanung.lerngruppeId)}
          fachObj={fach(lg(modalPlanung.lerngruppeId)?.fachId)}
          klasseObj={klass(modalPlanung.klasseId)}
          onClose={() => setModalPlanungId(null)}
          onStatus={(s) => setzeStatus(modalPlanung.id, s)}
          onNotiz={(n) => updatePlanung(modalPlanung.id, { notiz: n })}
          onRemove={() => removePlanung(modalPlanung.id)}
        />
      )}
    </div>
  );
}

// ---------- Klassen-Zeitleiste ----------
function KlasseAnsicht(props) {
  const {
    klasse, lehr, fach, meth, lg, lerngruppenFuerKlasse, planungenFuerKlasse, vorschlaege,
    onDragStartChip, onDragStartCard, onDropLane, onDropZurueckZuEmpfehlungen,
    onOeffne, manuell, setManuell, addManuell, methoden, lerngruppen,
  } = props;

  const faecherDerKlasse = lerngruppenFuerKlasse
    .slice()
    .sort((a, b) => a.fachId.localeCompare(b.fachId));

  const lerngruppenFuerManuell = manuell.methodeId
    ? lerngruppenFuerKlasse.filter((g) => meth(manuell.methodeId)?.faecherIds.includes(g.fachId))
    : [];

  return (
    <div>
      <div className="flex items-start justify-between flex-wrap gap-3 mb-6">
        <div>
          <h1 className="mc-display text-2xl font-semibold">
            Klasse {klasse.jahrgang}
            {klasse.buchstabe}
          </h1>
          <p className="text-sm" style={{ color: T.muted }}>
            Klassenlehrer: {[lehr(klasse.lehrer1)?.name, lehr(klasse.lehrer2)?.name].filter(Boolean).join(" · ") || "—"}
          </p>
        </div>
        <div className="flex gap-4 items-center text-xs" style={{ color: T.muted }}>
          <span className="flex items-center gap-1.5">
            <StatusDot status="ausstehend" /> ausstehend
          </span>
          <span className="flex items-center gap-1.5">
            <StatusDot status="erledigt" /> erledigt
          </span>
          <span className="flex items-center gap-1.5">
            <StatusDot status="ausgefallen" /> ausgefallen
          </span>
        </div>
      </div>

      {faecherDerKlasse.length === 0 ? (
        <div className="rounded-xl border p-12 text-center" style={{ borderColor: T.line, background: "white" }}>
          <p className="mc-display font-semibold mb-1">Noch keine Fächer hinterlegt</p>
          <p className="text-sm" style={{ color: T.muted }}>
            Für diese Klasse sind noch keine Lerngruppen angelegt (siehe Verwaltung).
          </p>
        </div>
      ) : (
        <div className="rounded-xl border overflow-hidden" style={{ borderColor: T.line, background: "white" }}>
          <div style={{ display: "grid", gridTemplateColumns: "150px repeat(4, 1fr)" }} className="border-b" >
            <div style={{ borderColor: T.line }} />
            {QUARTALE.map((q) => (
              <div key={q} className="px-4 py-2.5 border-l" style={{ borderColor: T.line }}>
                <div className="mc-display text-xs font-semibold uppercase tracking-wide" style={{ color: T.muted }}>
                  {q}. Quartal
                </div>
                <div className="mc-mono text-[10px]" style={{ color: T.muted }}>
                  {QUARTAL_SPANNE[q]}
                </div>
              </div>
            ))}
          </div>
          {faecherDerKlasse.map((g) => (
            <div
              key={g.id}
              style={{ display: "grid", gridTemplateColumns: "150px repeat(4, 1fr)" }}
              className="border-b last:border-b-0"
            >
              <div
                className="px-4 py-4 flex flex-col justify-center"
                style={{ background: T.paperAlt, borderColor: T.line, borderRight: `1px solid ${T.line}` }}
              >
                <span className="mc-display text-sm font-semibold">{fach(g.fachId)?.name}</span>
                <span className="mc-mono text-[10px]" style={{ color: T.muted }}>
                  {lehr(g.lehrerId)?.kuerzel || "—"}
                </span>
              </div>
              {QUARTALE.map((q) => {
                const karten = planungenFuerKlasse.filter((p) => {
                  const pg = lg(p.lerngruppeId);
                  return pg && pg.fachId === g.fachId && p.quartal === q;
                });
                return (
                  <div
                    key={q}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => onDropLane(e, klasse.id, g.fachId, q)}
                    className="px-3 py-3 border-l flex flex-wrap gap-1.5 content-start"
                    style={{ borderColor: T.line, minHeight: 60 }}
                  >
                    {karten.length === 0 ? (
                      <span className="text-[11px] italic" style={{ color: "#B4AF9F" }}>
                        – keine Methode –
                      </span>
                    ) : (
                      karten.map((p) => (
                        <MethodCard
                          key={p.id}
                          planung={p}
                          methode={meth(p.methodeId)}
                          onClick={() => onOeffne(p.id)}
                          onDragStart={(e) => onDragStartCard(e, p.id)}
                        />
                      ))
                    )}
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      )}

      {/* ---------------- Zuordnungs-Werkzeug ---------------- */}
      <div className="mt-8 grid gap-4" style={{ gridTemplateColumns: "1.1fr 1fr" }}>
        <div
          onDragOver={(e) => e.preventDefault()}
          onDrop={onDropZurueckZuEmpfehlungen}
          className="rounded-xl border p-4" style={{ borderColor: T.line, background: "white" }}>
          <div className="mc-display text-sm font-semibold mb-1">Empfehlungen für diese Klasse</div>
          <p className="text-xs mb-3" style={{ color: T.muted }}>
            Auf ein passendes Zeitfenster oben ziehen, um zuzuordnen – bereits platzierte Kärtchen lassen sich
            in ein anderes Fach oder Quartal, oder zurück hierher ziehen.
          </p>
          <div className="flex flex-wrap gap-2">
            {vorschlaege.length === 0 && (
              <p className="text-xs italic" style={{ color: T.muted }}>
                Keine weiteren Empfehlungen offen – alle passenden Methoden sind bereits zugeordnet.
              </p>
            )}
            {vorschlaege.map(({ methode, lerngruppe }) => (
              <div
                key={methode.id}
                draggable
                onDragStart={(e) => onDragStartChip(e, methode.id, lerngruppe.id)}
                className="cursor-grab active:cursor-grabbing rounded-lg px-3 py-2 text-xs"
                style={{ border: `1.5px dashed ${T.accent}`, background: T.accentSoft }}
                title="Ziehen, um einzuordnen"
              >
                <div className="font-semibold">{methode.name}</div>
                <div className="mc-mono" style={{ color: "#7A5518" }}>
                  {fach(lerngruppe.fachId)?.name} · empf. {methode.halbjahr}. HJ
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="rounded-xl border p-4" style={{ borderColor: T.line, background: "white" }}>
          <div className="mc-display text-sm font-semibold mb-3">Manuell zuordnen</div>
          <div className="flex flex-col gap-2">
            <Field label="Methode">
              <select
                className="border rounded px-2 py-1 text-sm"
                style={inputStyle()}
                value={manuell.methodeId}
                onChange={(e) => setManuell({ ...manuell, methodeId: e.target.value, lerngruppeId: "" })}
              >
                <option value="">wählen…</option>
                {methoden.map((m) => (
                  <option key={m.id} value={m.id}>
                    {m.name}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Fach">
              <select
                className="border rounded px-2 py-1 text-sm"
                style={inputStyle()}
                value={manuell.lerngruppeId}
                onChange={(e) => setManuell({ ...manuell, lerngruppeId: e.target.value })}
                disabled={!manuell.methodeId}
              >
                <option value="">wählen…</option>
                {lerngruppenFuerManuell.map((g) => (
                  <option key={g.id} value={g.id}>
                    {fach(g.fachId)?.name} ({lehr(g.lehrerId)?.kuerzel || "—"})
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Quartal">
              <select
                className="border rounded px-2 py-1 text-sm"
                style={inputStyle()}
                value={manuell.quartal}
                onChange={(e) => setManuell({ ...manuell, quartal: e.target.value })}
              >
                {QUARTALE.map((q) => (
                  <option key={q} value={q}>
                    {q}. Quartal
                  </option>
                ))}
              </select>
            </Field>
            <Button onClick={() => addManuell(klasse.id)}>+ Zuordnen</Button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Lehrer-zentrierte Ansicht ----------
function LehrerAnsicht({ lehrerObj, planungen, klass, meth, fach, lg, onOeffne }) {
  const [filter, setFilter] = useState({ klasseId: "", status: "" });
  const [sort, setSort] = useState({ feld: null, richtung: "auf" });

  if (!lehrerObj) {
    return (
      <div>
        <h1 className="mc-display text-2xl font-semibold mb-1">Meine Methoden</h1>
        <div className="rounded-xl border p-10 text-center mt-6" style={{ borderColor: T.line, background: "white" }}>
          <p className="text-sm" style={{ color: T.muted }}>
            Noch keine Lehrkräfte angelegt. Lehrkräfte lassen sich in der Verwaltung anlegen oder aus Untis
            importieren.
          </p>
        </div>
      </div>
    );
  }

  const klasseOptionen = [...new Map(planungen.map((p) => [p.klasseId, klass(p.klasseId)]).filter(([, k]) => k)).values()].sort(
    (a, b) => a.jahrgang - b.jahrgang || a.buchstabe.localeCompare(b.buchstabe)
  );
  const filterAktiv = filter.klasseId || filter.status;
  const gefiltert = planungen.filter(
    (p) => (!filter.klasseId || p.klasseId === filter.klasseId) && (!filter.status || p.status === filter.status)
  );
  const schluessel = (p, feld) => {
    if (feld === "methode") return meth(p.methodeId)?.name || "";
    if (feld === "klasse") {
      const k = klass(p.klasseId);
      return k ? `${String(k.jahrgang).padStart(2, "0")}${k.buchstabe}` : "";
    }
    if (feld === "fach") return fach(lg(p.lerngruppeId)?.fachId)?.name || "";
    if (feld === "quartal") return p.quartal;
    if (feld === "status") return p.status;
    return "";
  };
  const sortiert = sort.feld
    ? [...gefiltert].sort((a, b) => {
        const av = schluessel(a, sort.feld);
        const bv = schluessel(b, sort.feld);
        const cmp = typeof av === "number" ? av - bv : String(av).localeCompare(String(bv), "de");
        return sort.richtung === "auf" ? cmp : -cmp;
      })
    : gefiltert;
  const spaltenkopf = (feld, label) => (
    <button
      onClick={() => setSort((s) => (s.feld === feld ? { feld, richtung: s.richtung === "auf" ? "ab" : "auf" } : { feld, richtung: "auf" }))}
      className="text-left flex items-center gap-1 hover:underline"
    >
      {label}
      {sort.feld === feld && <span>{sort.richtung === "auf" ? "▲" : "▼"}</span>}
    </button>
  );

  return (
    <div>
      <h1 className="mc-display text-2xl font-semibold mb-1">Meine Methoden</h1>
      <p className="text-sm mb-4" style={{ color: T.muted }}>
        {lehrerObj.name || lehrerObj.kuerzel} · alle zugeordneten Methoden über alle Klassen und Kurse hinweg
      </p>
      {planungen.length === 0 ? (
        <div className="rounded-xl border p-10 text-center" style={{ borderColor: T.line, background: "white" }}>
          <p className="text-sm" style={{ color: T.muted }}>
            Für {lehrerObj.name || lehrerObj.kuerzel} sind aktuell keine Methoden zugeordnet.
          </p>
        </div>
      ) : (
        <>
          <div className="flex flex-wrap items-end gap-2 mb-2">
            <Field label="Klasse">
              <select className="border rounded px-2 py-1 text-xs" style={inputStyle()} value={filter.klasseId} onChange={(e) => setFilter({ ...filter, klasseId: e.target.value })}>
                <option value="">Alle</option>
                {klasseOptionen.map((k) => (
                  <option key={k.id} value={k.id}>{k.jahrgang}{k.buchstabe}</option>
                ))}
              </select>
            </Field>
            <Field label="Status">
              <select className="border rounded px-2 py-1 text-xs" style={inputStyle()} value={filter.status} onChange={(e) => setFilter({ ...filter, status: e.target.value })}>
                <option value="">Alle</option>
                <option value="ausstehend">ausstehend</option>
                <option value="erledigt">erledigt</option>
                <option value="ausgefallen">ausgefallen</option>
              </select>
            </Field>
            {filterAktiv && (
              <button onClick={() => setFilter({ klasseId: "", status: "" })} className="text-xs underline mb-1.5" style={{ color: T.muted }}>
                Filter zurücksetzen
              </button>
            )}
            <span className="text-xs mb-1.5 ml-auto" style={{ color: T.muted }}>
              {sortiert.length === planungen.length ? `${planungen.length} Einträge` : `${sortiert.length} von ${planungen.length} Einträgen`}
            </span>
          </div>
          <div className="rounded-xl border overflow-hidden" style={{ borderColor: T.line, background: "white" }}>
            <div className="flex items-center gap-3 px-4 py-2 text-[11px] uppercase tracking-wide font-medium border-b" style={{ borderColor: T.line, color: T.muted }}>
              <span className="w-40">{spaltenkopf("methode", "Methode")}</span>
              <span className="w-16">{spaltenkopf("klasse", "Klasse")}</span>
              <span className="w-28">{spaltenkopf("fach", "Fach")}</span>
              <span className="w-20">{spaltenkopf("quartal", "Quartal")}</span>
              <span className="ml-auto">{spaltenkopf("status", "Status")}</span>
            </div>
            <div className="divide-y">
              {sortiert.length === 0 ? (
                <div className="px-4 py-6 text-xs italic" style={{ color: T.muted }}>Keine Einträge für diese Filterauswahl.</div>
              ) : (
                sortiert.map((p) => {
                  const k = klass(p.klasseId);
                  const g = lg(p.lerngruppeId);
                  const m = meth(p.methodeId);
                  const tone = p.status === "erledigt" ? "success" : p.status === "ausgefallen" ? "cancelled" : "accent";
                  return (
                    <div
                      key={p.id}
                      onClick={() => onOeffne(p.id)}
                      className="flex items-center gap-3 px-4 py-3 text-sm cursor-pointer hover:bg-[#FAF9F5]"
                    >
                      <span className="font-medium w-40">{m.name}</span>
                      <span className="w-16" style={{ color: T.muted }}>
                        {k ? `${k.jahrgang}${k.buchstabe}` : "—"}
                      </span>
                      <span className="w-28" style={{ color: T.muted }}>
                        {fach(g?.fachId)?.name}
                      </span>
                      <span className="mc-mono text-xs w-20" style={{ color: T.muted }}>
                        {p.quartal}. Quartal
                      </span>
                      <span className="ml-auto">
                        <Badge tone={tone}>{STATUS_LABEL[p.status]}</Badge>
                      </span>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ---------- Detail-Modal ----------
// ---------- PDF-Vorschau für hochgeladene Materialien ----------
function istPdf(material) {
  return typeof material !== "string" && (material.name || "").toLowerCase().endsWith(".pdf");
}

function PdfVorschauModal({ material, onClose }) {
  return (
    <div
      className="fixed inset-0 z-[60] flex flex-col"
      style={{ background: "rgba(0,0,0,0.75)" }}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="flex items-center justify-between px-4 py-3 shrink-0" style={{ background: T.board, color: T.boardText }}>
        <span className="text-sm truncate">{material.name}</span>
        <div className="flex items-center gap-4 shrink-0">
          <a href={material.dataUrl} download={material.name} className="text-xs underline">
            Herunterladen
          </a>
          <button onClick={onClose} className="text-xl leading-none" aria-label="Vorschau schließen">
            ×
          </button>
        </div>
      </div>
      <iframe src={material.dataUrl} title={material.name} className="flex-1 w-full border-0" style={{ background: "white" }} />
    </div>
  );
}

function DetailModal({ planung, methode, gruppe, fachObj, klasseObj, onClose, onStatus, onNotiz, onRemove }) {
  const [pdfVorschau, setPdfVorschau] = useState(null);
  const tone = planung.status === "erledigt" ? "success" : planung.status === "ausgefallen" ? "cancelled" : "accent";
  return (
    <div
      className="fixed inset-0 flex items-center justify-center z-50"
      style={{ background: "rgba(27,58,63,0.45)" }}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="rounded-2xl p-6 w-[440px] max-w-[90vw] relative" style={{ background: "white", boxShadow: "0 20px 50px rgba(0,0,0,0.25)" }}>
        <button onClick={onClose} className="absolute top-3 right-4 text-xl" style={{ color: T.muted }}>
          ×
        </button>
        <Badge tone={tone}>{STATUS_LABEL[planung.status]}</Badge>
        <h2 className="mc-display text-xl font-semibold mt-2">{methode?.name}</h2>
        <p className="text-xs mb-3" style={{ color: T.muted }}>
          {fachObj?.name} · Klasse {klasseObj?.jahrgang}{klasseObj?.buchstabe} · {planung.quartal}. Quartal
        </p>

        <div className="mb-4">
          <div className="mc-display text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: T.muted }}>
            Beschreibung
          </div>
          <Beschreibungsfeld value={methode?.beschreibung} />
        </div>

        <div className="mb-4">
          <div className="mc-display text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: T.muted }}>
            Status
          </div>
          <div className="flex gap-1.5">
            {["ausstehend", "erledigt", "ausgefallen"].map((s) => (
              <button
                key={s}
                onClick={() => onStatus(s)}
                className="flex-1 text-xs rounded-lg py-1.5 border"
                style={
                  planung.status === s
                    ? { borderColor: T.accent, background: T.accentSoft, color: "#7A5518" }
                    : { borderColor: T.line, color: T.muted }
                }
              >
                {STATUS_LABEL[s]}
              </button>
            ))}
          </div>
          {planung.status === "erledigt" && planung.datum && (
            <p className="mc-mono text-xs mt-2" style={{ color: T.success }}>
              Durchgeführt am {new Date(planung.datum).toLocaleDateString("de-DE")}
            </p>
          )}
        </div>

        <div className="mb-4">
          <div className="mc-display text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: T.muted }}>
            Notiz
          </div>
          <textarea
            className="w-full text-sm rounded-lg p-2.5 border"
            style={{ borderColor: T.line, background: T.paperAlt, minHeight: 60 }}
            placeholder="Notiz zur Durchführung…"
            value={planung.notiz || ""}
            onChange={(e) => onNotiz(e.target.value)}
          />
        </div>

        <div className="mb-5">
          <div className="mc-display text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: T.muted }}>
            Materialien
          </div>
          {methode?.materialien?.length ? (
            <div className="flex flex-col gap-1.5">
              {methode.materialien.map((f, i) =>
                typeof f === "string" ? (
                  <div key={i} className="text-xs rounded-lg px-2.5 py-1.5 border flex items-center gap-2" style={{ borderColor: T.line, background: T.paperAlt, color: T.muted }}>
                    📄 {f}
                  </div>
                ) : istPdf(f) ? (
                  <button
                    key={i}
                    onClick={() => setPdfVorschau(f)}
                    className="text-xs rounded-lg px-2.5 py-1.5 border flex items-center gap-2 hover:underline text-left"
                    style={{ borderColor: T.line, background: T.paperAlt, color: T.ink }}
                  >
                    📄 {f.name} <span style={{ color: T.muted }}>· Vorschau</span>
                  </button>
                ) : (
                  <a
                    key={i}
                    href={f.dataUrl}
                    download={f.name}
                    className="text-xs rounded-lg px-2.5 py-1.5 border flex items-center gap-2 hover:underline"
                    style={{ borderColor: T.line, background: T.paperAlt, color: T.ink }}
                  >
                    📄 {f.name}
                  </a>
                )
              )}
            </div>
          ) : (
            <p className="text-xs italic" style={{ color: T.muted }}>
              Keine Materialien hinterlegt.
            </p>
          )}
        </div>

        {methode?.links?.length > 0 && (
          <div className="mb-5">
            <div className="mc-display text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: T.muted }}>
              Web-Links
            </div>
            <div className="flex flex-col gap-1.5">
              {methode.links.map((link, i) => (
                <a
                  key={i}
                  href={link.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-xs rounded-lg px-2.5 py-1.5 border flex items-center gap-2 hover:underline"
                  style={{ borderColor: T.line, background: T.paperAlt, color: T.ink }}
                >
                  🔗 {link.titel}
                </a>
              ))}
            </div>
          </div>
        )}

        <div className="flex justify-between items-center">
          <button onClick={onRemove} className="text-xs" style={{ color: T.danger }}>
            Zuordnung entfernen
          </button>
          <Button small tone="ghost" onClick={onClose}>
            Schließen
          </Button>
        </div>
      </div>
      {pdfVorschau && <PdfVorschauModal material={pdfVorschau} onClose={() => setPdfVorschau(null)} />}
    </div>
  );
}

// ---------- Verwaltung (Stammdaten + Methoden) ----------
function VerwaltungView(props) {
  const {
    faecher, lehrer, klassen, lerngruppen, methoden,
    setFaecher, setLehrer, setKlassen, setLerngruppen, setMethoden,
    planungen, setPlanungen,
    stammTab, setStammTab,
    neuFach, setNeuFach, addFach,
    neuLehrer, setNeuLehrer, addLehrer,
    updateKlasse, neuKlasse, setNeuKlasse, addKlasse,
    neuLg, setNeuLg, addLerngruppe, toggleKlasseInNeuLg, updateLerngruppe, removeLerngruppe,
    lerngruppenSicherung, undoLerngruppenAenderung, verwirfLerngruppenSicherung,
    neuMethode, setNeuMethode, addMethode, toggleInList, updateMethode, removeMethode, neuMethodeKey,
    datenquelle, migriereNachSupabase, migrationLaeuft, migrationFortschritt, migrationFehler,
  } = props;

  const fach = (id) => faecher.find((f) => f.id === id);

  // ----- Fächer: Suche + Sortierung -----
  const [faecherSuche, setFaecherSuche] = useState("");
  const [faecherSort, setFaecherSort] = useState({ feld: "name", richtung: "auf" });
  const faecherSpaltenkopf = (feld, label) => (
    <button onClick={() => setFaecherSort((s) => (s.feld === feld ? { feld, richtung: s.richtung === "auf" ? "ab" : "auf" } : { feld, richtung: "auf" }))} className="text-left flex items-center gap-1 hover:underline">
      {label}
      {faecherSort.feld === feld && <span>{faecherSort.richtung === "auf" ? "▲" : "▼"}</span>}
    </button>
  );
  const gefilterteFaecher = faecher.filter((f) => {
    const q = faecherSuche.trim().toLowerCase();
    return !q || f.name.toLowerCase().includes(q) || f.kuerzel.toLowerCase().includes(q);
  });
  const sortierteFaecher = [...gefilterteFaecher].sort((a, b) => {
    const cmp = (a[faecherSort.feld] || "").localeCompare(b[faecherSort.feld] || "", "de");
    return faecherSort.richtung === "auf" ? cmp : -cmp;
  });

  // ----- Lehrkräfte: Suche + Sortierung -----
  const [lehrerSuche, setLehrerSuche] = useState("");
  const [lehrerSort, setLehrerSort] = useState({ feld: "name", richtung: "auf" });
  const lehrerSpaltenkopf = (feld, label) => (
    <button onClick={() => setLehrerSort((s) => (s.feld === feld ? { feld, richtung: s.richtung === "auf" ? "ab" : "auf" } : { feld, richtung: "auf" }))} className="text-left flex items-center gap-1 hover:underline">
      {label}
      {lehrerSort.feld === feld && <span>{lehrerSort.richtung === "auf" ? "▲" : "▼"}</span>}
    </button>
  );
  const gefilterteLehrerListe = lehrer.filter((l) => {
    const q = lehrerSuche.trim().toLowerCase();
    return !q || (l.name || "").toLowerCase().includes(q) || (l.email || "").toLowerCase().includes(q) || l.kuerzel.toLowerCase().includes(q);
  });
  const sortierteLehrerListe = [...gefilterteLehrerListe].sort((a, b) => {
    const cmp = (a[lehrerSort.feld] || "").localeCompare(b[lehrerSort.feld] || "", "de");
    return lehrerSort.richtung === "auf" ? cmp : -cmp;
  });

  // ----- Methoden: Filter (Jahrgang, Fach) + Sortierung -----
  const [methodenFilter, setMethodenFilter] = useState({ jahrgang: "", fachId: "" });
  const [editModal, setEditModal] = useState(null); // { methodeId } | { neu: true } | null
  const [pdfVorschau, setPdfVorschau] = useState(null);
  const [methodeLoeschenBestaetigen, setMethodeLoeschenBestaetigen] = useState(null); // methodeId oder null

  // Materialien (Arbeitsblätter etc.) je Methode: als Base64-Datei-URL im Zustand gehalten,
  // dadurch ohne eigenen Server herunterladbar. Nicht für sehr große Dateien gedacht.
  const materialHinzufuegen = (methodeId, dateien) => {
    Promise.all(
      Array.from(dateien).map(
        (datei) =>
          new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve({ name: datei.name, dataUrl: reader.result });
            reader.readAsDataURL(datei);
          })
      )
    ).then((neue) => {
      const aktuell = methoden.find((x) => x.id === methodeId);
      updateMethode(methodeId, { materialien: [...(aktuell?.materialien || []), ...neue] });
    });
  };
  const materialEntfernen = (methodeId, index) => {
    const aktuell = methoden.find((x) => x.id === methodeId);
    updateMethode(methodeId, { materialien: (aktuell?.materialien || []).filter((_, i) => i !== index) });
  };

  // Web-Links je Methode: { titel, url }. neuerLink hält den Entwurf pro Methode
  // (mehrere Karten gleichzeitig, jede mit eigenem kleinen Eingabeformular).
  const [neuerLink, setNeuerLink] = useState({});
  const linkHinzufuegen = (methodeId) => {
    const entwurf = neuerLink[methodeId];
    if (!entwurf?.url?.trim()) return;
    const url = /^https?:\/\//i.test(entwurf.url.trim()) ? entwurf.url.trim() : `https://${entwurf.url.trim()}`;
    const aktuell = methoden.find((x) => x.id === methodeId);
    updateMethode(methodeId, { links: [...(aktuell?.links || []), { titel: entwurf.titel?.trim() || url, url }] });
    setNeuerLink({ ...neuerLink, [methodeId]: { titel: "", url: "" } });
  };
  const linkEntfernen = (methodeId, index) => {
    const aktuell = methoden.find((x) => x.id === methodeId);
    updateMethode(methodeId, { links: (aktuell?.links || []).filter((_, i) => i !== index) });
  };

  // ----- Methoden-Export/Import (JSON) -----
  const methodeExportieren = (m) => {
    jsonDateiHerunterladen({ typ: "methodencurriculum-methode", version: 1, methode: methodeZuExportobjekt(m, faecher) }, `${slug(m.name)}.methode.json`);
  };
  const alleMethodenExportieren = () => {
    jsonDateiHerunterladen(
      { typ: "methodencurriculum-methoden-satz", version: 1, methoden: methoden.map((m) => methodeZuExportobjekt(m, faecher)) },
      "methoden-satz.json"
    );
  };
  const [methodenImportFehler, setMethodenImportFehler] = useState("");
  const methodenImportieren = (e) => {
    const datei = e.target.files[0];
    if (!datei) return;
    setMethodenImportFehler("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const payload = JSON.parse(reader.result);
        let objekte, ersetztKompletteListe;
        if (payload?.typ === "methodencurriculum-methode" && payload.methode) {
          objekte = [payload.methode];
          ersetztKompletteListe = false;
        } else if (payload?.typ === "methodencurriculum-methoden-satz" && Array.isArray(payload.methoden)) {
          objekte = payload.methoden;
          ersetztKompletteListe = true;
        } else {
          setMethodenImportFehler("Diese Datei sieht nicht wie ein Methoden-Export dieser App aus.");
          return;
        }
        const { faecherPool, neueMethoden } = bauMethodenAusExport(objekte, faecher);
        setFaecher(faecherPool);
        setMethoden(ersetztKompletteListe ? neueMethoden : [...methoden, ...neueMethoden]);
      } catch (err) {
        setMethodenImportFehler("Datei konnte nicht gelesen werden (kein gültiges JSON).");
      }
    };
    reader.readAsText(datei);
    e.target.value = "";
  };

  const [methodenSort, setMethodenSort] = useState({ feld: "name", richtung: "auf" });
  const methodenSpaltenkopf = (feld, label) => (
    <button onClick={() => setMethodenSort((s) => (s.feld === feld ? { feld, richtung: s.richtung === "auf" ? "ab" : "auf" } : { feld, richtung: "auf" }))} className="text-xs flex items-center gap-1 hover:underline" style={{ color: T.muted }}>
      {label}
      {methodenSort.feld === feld && <span>{methodenSort.richtung === "auf" ? "▲" : "▼"}</span>}
    </button>
  );
  const methodenFilterAktiv = methodenFilter.jahrgang || methodenFilter.fachId;
  const gefilterteMethoden = methoden.filter(
    (m) =>
      (!methodenFilter.jahrgang || m.jahrgaenge.includes(Number(methodenFilter.jahrgang))) &&
      (!methodenFilter.fachId || m.faecherIds.includes(methodenFilter.fachId))
  );
  const sortierteMethoden = [...gefilterteMethoden].sort((a, b) => {
    let cmp;
    if (methodenSort.feld === "jahrgang") {
      cmp = (Math.min(...(a.jahrgaenge.length ? a.jahrgaenge : [99]))) - (Math.min(...(b.jahrgaenge.length ? b.jahrgaenge : [99])));
    } else {
      cmp = a.name.localeCompare(b.name, "de");
    }
    return methodenSort.richtung === "auf" ? cmp : -cmp;
  });

  const [lgFilter, setLgFilter] = useState({ klasseId: "", fachId: "", lehrerId: "" });
  const lgFilterOptionenKlassen = klassen
    .filter((k) => lerngruppen.some((g) => g.klassenIds.includes(k.id)))
    .sort((a, b) => a.jahrgang - b.jahrgang || a.buchstabe.localeCompare(b.buchstabe));
  const lgFilterOptionenFaecher = faecher
    .filter((f) => lerngruppen.some((g) => g.fachId === f.id))
    .sort((a, b) => a.name.localeCompare(b.name, "de"));
  const lgFilterOptionenLehrer = lehrer
    .filter((l) => lerngruppen.some((g) => g.lehrerId === l.id))
    .sort((a, b) => a.kuerzel.localeCompare(b.kuerzel, "de"));
  const lgFilterAktiv = lgFilter.klasseId || lgFilter.fachId || lgFilter.lehrerId;
  const gefilterteLerngruppen = lerngruppen.filter(
    (g) =>
      (!lgFilter.klasseId || g.klassenIds.includes(lgFilter.klasseId)) &&
      (!lgFilter.fachId || g.fachId === lgFilter.fachId) &&
      (!lgFilter.lehrerId || g.lehrerId === lgFilter.lehrerId)
  );

  const [lgSortierung, setLgSortierung] = useState({ feld: null, richtung: "auf" });
  const klickLgSortierung = (feld) => {
    setLgSortierung((s) =>
      s.feld === feld ? { feld, richtung: s.richtung === "auf" ? "ab" : "auf" } : { feld, richtung: "auf" }
    );
  };
  const lgSortierSchluessel = (g, feld) => {
    if (feld === "klasse") return g.bezeichnung || "";
    if (feld === "fach") return fach(g.fachId)?.name || "";
    if (feld === "lehrer") return lehrer.find((l) => l.id === g.lehrerId)?.kuerzel || "";
    return "";
  };
  const sortierteLerngruppen = lgSortierung.feld
    ? [...gefilterteLerngruppen].sort((a, b) => {
        const cmp = lgSortierSchluessel(a, lgSortierung.feld).localeCompare(lgSortierSchluessel(b, lgSortierung.feld), "de");
        return lgSortierung.richtung === "auf" ? cmp : -cmp;
      })
    : gefilterteLerngruppen;
  const lgSpaltenkopf = (feld, label) => (
    <button
      onClick={() => klickLgSortierung(feld)}
      className="text-left flex items-center gap-1 hover:underline"
    >
      {label}
      {lgSortierung.feld === feld && <span>{lgSortierung.richtung === "auf" ? "▲" : "▼"}</span>}
    </button>
  );

  const exportiereLerngruppenCSV = () => {
    const zeilen = lerngruppen.map((g) => ({
      Fach: fach(g.fachId)?.name || "",
      FachKuerzel: fach(g.fachId)?.kuerzel || "",
      Bezeichnung: g.bezeichnung,
      Jahrgang: g.jahrgang,
      LehrerKuerzel: lehrer.find((l) => l.id === g.lehrerId)?.kuerzel || "",
      Klassen: g.klassenIds
        .map((id) => klassen.find((k) => k.id === id))
        .filter(Boolean)
        .map((k) => `${k.jahrgang}${k.buchstabe}`)
        .join("~"),
    }));
    const csv = Papa.unparse(zeilen, { delimiter: ";" });
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "lerngruppen.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <h1 className="mc-display text-2xl font-semibold mb-1">Verwaltung</h1>
      <p className="text-sm mb-5" style={{ color: T.muted }}>
        Stammdaten der Plattform – künftig teilweise aus IServ/Untis importierbar.
      </p>

      {datenquelle === "lokal" && (
        <div className="rounded-xl border p-4 mb-5" style={{ borderColor: T.accent, background: T.accentSoft }}>
          <div className="mc-display text-sm font-semibold mb-1" style={{ color: "#7A5518" }}>
            Noch nicht mit Supabase verbunden
          </div>
          <p className="text-xs mb-3" style={{ color: "#7A5518" }}>
            Aktuell läuft die App noch mit dem lokalen Ausgangszustand (nichts wird beim Neuladen gespeichert). Mit
            einem Klick lässt sich der komplette aktuelle Stand einmalig nach Supabase hochladen – danach lädt die
            App bei jedem Start automatisch von dort und alle Änderungen werden dauerhaft gespeichert.
          </p>
          <Button tone="accent" onClick={migriereNachSupabase} disabled={migrationLaeuft}>
            {migrationLaeuft ? `Übertrage… ${migrationFortschritt}` : "Jetzt einmalig nach Supabase übernehmen"}
          </Button>
          {migrationFehler && (
            <p className="text-xs mt-2" style={{ color: T.danger }}>
              Fehlgeschlagen: {migrationFehler}
            </p>
          )}
        </div>
      )}

      <div className="flex gap-1 mb-5 flex-wrap">
        {[
          ["faecher", "Fächer"],
          ["lehrer", "Lehrkräfte"],
          ["klassen", "Klassen"],
          ["lerngruppen", "Lerngruppen"],
          ["methoden", "Methoden"],
          ["import", "Untis-Import"],
        ].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setStammTab(key)}
            className="mc-mono text-xs px-3 py-1 rounded-full border uppercase tracking-wide"
            style={
              stammTab === key
                ? { background: T.accentSoft, borderColor: T.accent, color: "#7A5518" }
                : { background: "transparent", borderColor: T.line, color: T.muted }
            }
          >
            {label}
          </button>
        ))}
      </div>

      {stammTab === "faecher" && (
        <div className="max-w-md">
          <div className="flex items-center gap-2 mb-2">
            <input
              className="border rounded px-2 py-1 text-sm flex-1"
              style={inputStyle()}
              placeholder="Suche nach Name oder Kürzel…"
              value={faecherSuche}
              onChange={(e) => setFaecherSuche(e.target.value)}
            />
            <span className="text-xs shrink-0" style={{ color: T.muted }}>
              {gefilterteFaecher.length} von {faecher.length}
            </span>
          </div>
          <div className="rounded border bg-white overflow-hidden" style={{ borderColor: T.line }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 80px" }} className="px-3 py-1.5 text-[11px] uppercase tracking-wide font-medium border-b">
              <span style={{ color: T.muted }}>{faecherSpaltenkopf("name", "Name")}</span>
              <span style={{ color: T.muted }}>{faecherSpaltenkopf("kuerzel", "Kürzel")}</span>
            </div>
            <div className="divide-y">
              {sortierteFaecher.length === 0 ? (
                <div className="px-3 py-4 text-xs italic" style={{ color: T.muted }}>Keine Treffer.</div>
              ) : (
                sortierteFaecher.map((f) => (
                  <div key={f.id} style={{ display: "grid", gridTemplateColumns: "1fr 80px" }} className="px-3 py-2 text-sm">
                    <span>{f.name}</span>
                    <span className="mc-mono text-xs" style={{ color: T.muted }}>{f.kuerzel}</span>
                  </div>
                ))
              )}
            </div>
          </div>
          <div className="flex gap-2 mt-3">
            <input className="border rounded px-2 py-1 text-sm flex-1" style={inputStyle()} placeholder="Name" value={neuFach.name} onChange={(e) => setNeuFach({ ...neuFach, name: e.target.value })} />
            <input className="border rounded px-2 py-1 text-sm w-20" style={inputStyle()} placeholder="Kürzel" value={neuFach.kuerzel} onChange={(e) => setNeuFach({ ...neuFach, kuerzel: e.target.value })} />
            <Button onClick={addFach}>+ Hinzufügen</Button>
          </div>
        </div>
      )}

      {stammTab === "lehrer" && (
        <div className="max-w-xl">
          <div className="flex items-center gap-2 mb-2">
            <input
              className="border rounded px-2 py-1 text-sm flex-1"
              style={inputStyle()}
              placeholder="Suche nach Name, E-Mail oder Kürzel…"
              value={lehrerSuche}
              onChange={(e) => setLehrerSuche(e.target.value)}
            />
            <span className="text-xs shrink-0" style={{ color: T.muted }}>
              {gefilterteLehrerListe.length} von {lehrer.length}
            </span>
          </div>
          <div className="rounded border bg-white overflow-hidden" style={{ borderColor: T.line }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 70px" }} className="px-3 py-1.5 text-[11px] uppercase tracking-wide font-medium border-b">
              <span style={{ color: T.muted }}>{lehrerSpaltenkopf("name", "Name")}</span>
              <span style={{ color: T.muted }}>{lehrerSpaltenkopf("email", "E-Mail")}</span>
              <span style={{ color: T.muted }}>{lehrerSpaltenkopf("kuerzel", "Kürzel")}</span>
            </div>
            <div className="divide-y">
              {sortierteLehrerListe.length === 0 ? (
                <div className="px-3 py-4 text-xs italic" style={{ color: T.muted }}>Keine Treffer.</div>
              ) : (
                sortierteLehrerListe.map((l) => (
                  <div key={l.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 70px" }} className="px-3 py-2 text-sm">
                    <span className="truncate">{l.name || "—"}</span>
                    <span className="truncate" style={{ color: T.muted }}>{l.email || "—"}</span>
                    <span className="mc-mono text-xs" style={{ color: T.muted }}>{l.kuerzel}</span>
                  </div>
                ))
              )}
            </div>
          </div>
          <div className="flex gap-2 mt-3 flex-wrap">
            <input className="border rounded px-2 py-1 text-sm flex-1" style={inputStyle()} placeholder="Name" value={neuLehrer.name} onChange={(e) => setNeuLehrer({ ...neuLehrer, name: e.target.value })} />
            <input className="border rounded px-2 py-1 text-sm w-48" style={inputStyle()} placeholder="E-Mail" value={neuLehrer.email} onChange={(e) => setNeuLehrer({ ...neuLehrer, email: e.target.value })} />
            <input className="border rounded px-2 py-1 text-sm w-20" style={inputStyle()} placeholder="Kürzel (3)" maxLength={3} value={neuLehrer.kuerzel} onChange={(e) => setNeuLehrer({ ...neuLehrer, kuerzel: e.target.value.toUpperCase() })} />
            <Button onClick={addLehrer}>+ Hinzufügen</Button>
          </div>
        </div>
      )}

      {stammTab === "klassen" && (
        <div>
          <div className="grid grid-cols-3 gap-3">
            {JAHRGAENGE.map((jg) => (
              <div key={jg}>
                <div className="mc-mono text-xs mb-1 uppercase tracking-wide" style={{ color: T.muted }}>Jahrgang {jg}</div>
                <div className="space-y-2">
                  {klassen.filter((k) => k.jahrgang === jg).map((k) => (
                    <div key={k.id} className="rounded border p-2 text-sm bg-white" style={{ borderColor: T.line }}>
                      <div className="font-medium mb-1">{k.jahrgang}{k.buchstabe}</div>
                      <div className="flex gap-1 text-xs">
                        <select className="border rounded px-1 py-0.5 flex-1" style={inputStyle()} value={k.lehrer1 || ""} onChange={(e) => updateKlasse(k.id, { lehrer1: e.target.value })}>
                          <option value="">Klassenlehrer 1</option>
                          {lehrer.map((l) => <option key={l.id} value={l.id}>{l.kuerzel}</option>)}
                        </select>
                        <select className="border rounded px-1 py-0.5 flex-1" style={inputStyle()} value={k.lehrer2 || ""} onChange={(e) => updateKlasse(k.id, { lehrer2: e.target.value })}>
                          <option value="">Klassenlehrer 2</option>
                          {lehrer.map((l) => <option key={l.id} value={l.id}>{l.kuerzel}</option>)}
                        </select>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <div className="flex gap-2 mt-4 items-end">
            <Field label="Jahrgang">
              <select className="border rounded px-2 py-1 text-sm" style={inputStyle()} value={neuKlasse.jahrgang} onChange={(e) => setNeuKlasse({ ...neuKlasse, jahrgang: e.target.value })}>
                {JAHRGAENGE.map((j) => <option key={j} value={j}>{j}</option>)}
              </select>
            </Field>
            <Field label="Buchstabe">
              <input className="border rounded px-2 py-1 text-sm w-16" style={inputStyle()} value={neuKlasse.buchstabe} onChange={(e) => setNeuKlasse({ ...neuKlasse, buchstabe: e.target.value })} />
            </Field>
            <Button onClick={addKlasse}>+ Klasse anlegen</Button>
          </div>
        </div>
      )}

      {stammTab === "lerngruppen" && (
        <div>
          <div className="flex items-center justify-between mb-2 max-w-2xl">
            <div className="text-xs font-medium" style={{ color: T.muted }}>
              {gefilterteLerngruppen.length === lerngruppen.length
                ? `${lerngruppen.length} Lerngruppen`
                : `${gefilterteLerngruppen.length} von ${lerngruppen.length} Lerngruppen`}
            </div>
            <Button small tone="ghost" onClick={exportiereLerngruppenCSV}>
              Als CSV exportieren
            </Button>
          </div>
          <div className="flex flex-wrap items-end gap-2 mb-2 max-w-2xl">
            <Field label="Klasse">
              <select
                className="border rounded px-2 py-1 text-xs"
                style={inputStyle()}
                value={lgFilter.klasseId}
                onChange={(e) => setLgFilter({ ...lgFilter, klasseId: e.target.value })}
              >
                <option value="">Alle</option>
                {lgFilterOptionenKlassen.map((k) => (
                  <option key={k.id} value={k.id}>
                    {k.jahrgang}{k.buchstabe}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Fach">
              <select
                className="border rounded px-2 py-1 text-xs"
                style={inputStyle()}
                value={lgFilter.fachId}
                onChange={(e) => setLgFilter({ ...lgFilter, fachId: e.target.value })}
              >
                <option value="">Alle</option>
                {lgFilterOptionenFaecher.map((f) => (
                  <option key={f.id} value={f.id}>
                    {f.name}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Lehrer">
              <select
                className="border rounded px-2 py-1 text-xs"
                style={inputStyle()}
                value={lgFilter.lehrerId}
                onChange={(e) => setLgFilter({ ...lgFilter, lehrerId: e.target.value })}
              >
                <option value="">Alle</option>
                {lgFilterOptionenLehrer.map((l) => (
                  <option key={l.id} value={l.id}>
                    {l.kuerzel}
                  </option>
                ))}
              </select>
            </Field>
            {lgFilterAktiv && (
              <button
                onClick={() => setLgFilter({ klasseId: "", fachId: "", lehrerId: "" })}
                className="text-xs underline mb-1.5"
                style={{ color: T.muted }}
              >
                Filter zurücksetzen
              </button>
            )}
          </div>
          {lerngruppenSicherung && (
            <div
              className="flex items-center justify-between text-xs rounded px-3 py-2 mb-2 max-w-2xl"
              style={{ background: T.accentSoft, color: "#7A5518" }}
            >
              <span>{lerngruppenSicherung.beschreibung}</span>
              <span className="flex items-center gap-3 shrink-0">
                <button onClick={undoLerngruppenAenderung} className="font-medium underline">
                  Rückgängig
                </button>
                <button onClick={verwirfLerngruppenSicherung} aria-label="Hinweis schließen" style={{ color: "#7A5518" }}>
                  ×
                </button>
              </span>
            </div>
          )}
          <div className="rounded border bg-white max-w-2xl overflow-hidden" style={{ borderColor: T.line }}>
            <div
              style={{ display: "grid", gridTemplateColumns: "110px 1fr 90px 90px", borderColor: T.line }}
              className="px-3 py-1.5 text-[11px] uppercase tracking-wide font-medium border-b"
            >
              <span style={{ color: T.muted }}>{lgSpaltenkopf("klasse", "Klasse")}</span>
              <span style={{ color: T.muted }}>{lgSpaltenkopf("fach", "Fach")}</span>
              <span style={{ color: T.muted }}>{lgSpaltenkopf("lehrer", "Lehrer")}</span>
              <span></span>
            </div>
            <div className="divide-y">
              {sortierteLerngruppen.length === 0 ? (
                <div className="px-3 py-4 text-xs italic" style={{ color: T.muted }}>
                  Keine Lerngruppen für diese Filterauswahl.
                </div>
              ) : (
                sortierteLerngruppen.map((g) => (
                  <LerngruppeZeile key={g.id} g={g} faecher={faecher} lehrer={lehrer} klassen={klassen} updateLerngruppe={updateLerngruppe} removeLerngruppe={removeLerngruppe} />
                ))
              )}
            </div>
          </div>
          <div className="border rounded p-3 mt-3 max-w-2xl bg-white" style={{ borderColor: T.line }}>
            <div className="text-xs font-medium mb-2" style={{ color: T.muted }}>Neue Lerngruppe / neuer Kurs</div>
            <div className="flex gap-2 flex-wrap items-end">
              <Field label="Fach">
                <select className="border rounded px-2 py-1 text-sm" style={inputStyle()} value={neuLg.fachId} onChange={(e) => setNeuLg({ ...neuLg, fachId: e.target.value })}>
                  {faecher.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
                </select>
              </Field>
              <Field label="Bezeichnung">
                <input className="border rounded px-2 py-1 text-sm w-32" style={inputStyle()} value={neuLg.bezeichnung} onChange={(e) => setNeuLg({ ...neuLg, bezeichnung: e.target.value })} />
              </Field>
              <Field label="Jahrgang">
                <select className="border rounded px-2 py-1 text-sm" style={inputStyle()} value={neuLg.jahrgang} onChange={(e) => setNeuLg({ ...neuLg, jahrgang: e.target.value })}>
                  {JAHRGAENGE.map((j) => <option key={j} value={j}>{j}</option>)}
                </select>
              </Field>
              <Field label="Fachlehrer">
                <select className="border rounded px-2 py-1 text-sm" style={inputStyle()} value={neuLg.lehrerId} onChange={(e) => setNeuLg({ ...neuLg, lehrerId: e.target.value })}>
                  <option value="">wählen…</option>
                  {lehrer.map((l) => <option key={l.id} value={l.id}>{l.kuerzel}</option>)}
                </select>
              </Field>
              <Button onClick={addLerngruppe}>+ Anlegen</Button>
            </div>
            <div className="mt-2 flex flex-wrap gap-1">
              {klassen.filter((k) => k.jahrgang === Number(neuLg.jahrgang)).map((k) => (
                <button key={k.id} onClick={() => toggleKlasseInNeuLg(k.id)} className="text-xs px-2 py-1 rounded border" style={neuLg.klassenIds.includes(k.id) ? { background: T.accentSoft, borderColor: T.accent } : { borderColor: T.line }}>
                  {k.jahrgang}{k.buchstabe}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {stammTab === "methoden" && (
        <div>
          <div className="flex flex-wrap items-end gap-2 mb-3">
            <Field label="Jahrgang">
              <select
                className="border rounded px-2 py-1 text-xs"
                style={inputStyle()}
                value={methodenFilter.jahrgang}
                onChange={(e) => setMethodenFilter({ ...methodenFilter, jahrgang: e.target.value })}
              >
                <option value="">Alle</option>
                {JAHRGAENGE.map((jg) => (
                  <option key={jg} value={jg}>{jg}</option>
                ))}
              </select>
            </Field>
            <Field label="Fach">
              <select
                className="border rounded px-2 py-1 text-xs"
                style={inputStyle()}
                value={methodenFilter.fachId}
                onChange={(e) => setMethodenFilter({ ...methodenFilter, fachId: e.target.value })}
              >
                <option value="">Alle</option>
                {faecher.map((f) => (
                  <option key={f.id} value={f.id}>{f.name}</option>
                ))}
              </select>
            </Field>
            {methodenFilterAktiv && (
              <button onClick={() => setMethodenFilter({ jahrgang: "", fachId: "" })} className="text-xs underline mb-1.5" style={{ color: T.muted }}>
                Filter zurücksetzen
              </button>
            )}
            <span className="text-xs mb-1.5 ml-auto" style={{ color: T.muted }}>
              {gefilterteMethoden.length === methoden.length ? `${methoden.length} Methoden` : `${gefilterteMethoden.length} von ${methoden.length} Methoden`}
            </span>
            <span className="flex items-center gap-2 mb-1.5">
              <span className="text-xs" style={{ color: T.muted }}>Sortieren:</span>
              {methodenSpaltenkopf("name", "Name")}
              {methodenSpaltenkopf("jahrgang", "Jahrgang")}
            </span>
          </div>
          <div className="flex flex-wrap items-center gap-2 mb-3 pb-3 border-b" style={{ borderColor: T.line }}>
            <Button small tone="ghost" onClick={alleMethodenExportieren}>
              Alle exportieren
            </Button>
            <label className="text-xs px-2 py-1 rounded border cursor-pointer" style={{ borderColor: T.line, color: T.ink }}>
              Importieren…
              <input type="file" accept=".json" className="hidden" onChange={methodenImportieren} />
            </label>
            <span className="text-xs" style={{ color: T.muted }}>
              (einzelne Methode wird ergänzt, ein kompletter Satz ersetzt die aktuelle Liste)
            </span>
            {methodenImportFehler && (
              <span className="text-xs" style={{ color: T.danger }}>
                {methodenImportFehler}
              </span>
            )}
          </div>
          <div className="grid gap-3 mb-5" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))" }}>
            {sortierteMethoden.length === 0 ? (
              <p className="text-xs italic" style={{ color: T.muted }}>Keine Methoden für diese Filterauswahl.</p>
            ) : (
              sortierteMethoden.map((m) => (
              <div key={m.id} className="rounded border p-3 bg-white" style={{ borderColor: T.line }}>
                <div className="flex items-start justify-between gap-2 mb-1.5">
                  <input
                    className="font-medium text-sm border rounded px-1.5 py-0.5 flex-1 min-w-0"
                    style={inputStyle()}
                    value={m.name}
                    onChange={(e) => updateMethode(m.id, { name: e.target.value })}
                  />
                  <button
                    onClick={() => setEditModal({ methodeId: m.id })}
                    title="Beschreibung bearbeiten"
                    className="text-xs shrink-0"
                    style={{ color: T.accent }}
                  >
                    ✎ Bearbeiten
                  </button>
                  <button
                    onClick={() => methodeExportieren(m)}
                    title="Diese Methode als Datei exportieren"
                    className="text-xs shrink-0"
                    style={{ color: T.muted }}
                  >
                    ⬇ Export
                  </button>
                  {methodeLoeschenBestaetigen === m.id ? (
                    <span className="flex items-center gap-1.5 shrink-0">
                      <button
                        onClick={() => {
                          removeMethode(m.id);
                          setMethodeLoeschenBestaetigen(null);
                        }}
                        className="text-xs font-medium"
                        style={{ color: T.danger }}
                      >
                        Ja, löschen
                      </button>
                      <button onClick={() => setMethodeLoeschenBestaetigen(null)} className="text-xs" style={{ color: T.muted }}>
                        Abbr.
                      </button>
                    </span>
                  ) : (
                    <button
                      onClick={() => setMethodeLoeschenBestaetigen(m.id)}
                      title="Methode löschen"
                      className="text-xs shrink-0"
                      style={{ color: T.danger }}
                    >
                      Löschen
                    </button>
                  )}
                </div>
                <div className="mb-2">
                  <Beschreibungsfeld value={m.beschreibung} />
                </div>
                <div className="mb-2">
                  <div className="text-[11px] font-medium mb-1" style={{ color: T.muted }}>
                    Jahrgangsstufen
                  </div>
                  <div className="flex flex-wrap gap-1 mb-2">
                    {JAHRGAENGE.map((jg) => (
                      <button
                        key={jg}
                        onClick={() => updateMethode(m.id, { jahrgaenge: toggleInList(m.jahrgaenge, jg) })}
                        className="text-xs px-2 py-1 rounded border"
                        style={m.jahrgaenge.includes(jg) ? { background: T.accentSoft, borderColor: T.accent } : { borderColor: T.line }}
                      >
                        {jg}
                      </button>
                    ))}
                  </div>
                  <div className="text-[11px] font-medium mb-1" style={{ color: T.muted }}>
                    Empfohlene Fächer (Reihenfolge = Priorität)
                  </div>
                  {faecher.length === 0 ? (
                    <p className="text-xs italic" style={{ color: T.muted }}>
                      Noch keine Fächer angelegt.
                    </p>
                  ) : (
                    <div className="flex flex-wrap gap-1">
                      {faecher.map((f) => (
                        <button
                          key={f.id}
                          onClick={() => updateMethode(m.id, { faecherIds: toggleInList(m.faecherIds, f.id) })}
                          className="text-xs px-2 py-1 rounded border"
                          style={m.faecherIds.includes(f.id) ? { background: T.accentSoft, borderColor: T.accent } : { borderColor: T.line }}
                        >
                          {f.name}
                          {m.faecherIds.includes(f.id) && ` (${m.faecherIds.indexOf(f.id) + 1})`}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div className="flex flex-wrap items-center gap-1 mb-2">
                  <span className="text-[11px] font-medium mr-1" style={{ color: T.muted }}>
                    Halbjahr:
                  </span>
                  {[1, 2].map((h) => (
                    <button
                      key={h}
                      onClick={() => updateMethode(m.id, { halbjahr: h })}
                      className="text-xs px-2 py-1 rounded border"
                      style={m.halbjahr === h ? { background: T.ink, color: T.paper, borderColor: T.ink } : { borderColor: T.line }}
                    >
                      {h}. HJ
                    </button>
                  ))}
                </div>
                <div>
                  <div className="text-[11px] font-medium mb-1" style={{ color: T.muted }}>
                    Materialien
                  </div>
                  <div className="flex flex-col gap-1 mb-1.5">
                    {(m.materialien || []).map((mat, i) => (
                      <div key={i} className="flex items-center justify-between text-xs rounded border px-2 py-1" style={{ borderColor: T.line, background: T.paperAlt }}>
                        {istPdf(mat) ? (
                          <button onClick={() => setPdfVorschau(mat)} className="truncate text-left hover:underline" style={{ color: T.ink }}>
                            📄 {mat.name} <span style={{ color: T.muted }}>· Vorschau</span>
                          </button>
                        ) : (
                          <span className="truncate" style={{ color: T.muted }}>
                            📄 {typeof mat === "string" ? mat : mat.name}
                          </span>
                        )}
                        <button onClick={() => materialEntfernen(m.id, i)} className="shrink-0 ml-2" style={{ color: T.danger }}>
                          ×
                        </button>
                      </div>
                    ))}
                  </div>
                  <label className="text-xs cursor-pointer" style={{ color: T.accent }}>
                    + Material hochladen
                    <input
                      type="file"
                      multiple
                      className="hidden"
                      onChange={(e) => {
                        if (e.target.files.length) materialHinzufuegen(m.id, e.target.files);
                        e.target.value = "";
                      }}
                    />
                  </label>
                </div>
                <div className="mt-2">
                  <div className="text-[11px] font-medium mb-1" style={{ color: T.muted }}>
                    Web-Links
                  </div>
                  <div className="flex flex-col gap-1 mb-1.5">
                    {(m.links || []).map((link, i) => (
                      <div key={i} className="flex items-center justify-between text-xs rounded border px-2 py-1" style={{ borderColor: T.line, background: T.paperAlt }}>
                        <a href={link.url} target="_blank" rel="noopener noreferrer" className="truncate hover:underline" style={{ color: T.ink }}>
                          🔗 {link.titel}
                        </a>
                        <button onClick={() => linkEntfernen(m.id, i)} className="shrink-0 ml-2" style={{ color: T.danger }}>
                          ×
                        </button>
                      </div>
                    ))}
                  </div>
                  <div className="flex gap-1">
                    <input
                      className="border rounded px-1.5 py-1 text-xs flex-1 min-w-0"
                      style={inputStyle()}
                      placeholder="Titel (optional)"
                      value={neuerLink[m.id]?.titel || ""}
                      onChange={(e) => setNeuerLink({ ...neuerLink, [m.id]: { ...neuerLink[m.id], titel: e.target.value } })}
                    />
                    <input
                      className="border rounded px-1.5 py-1 text-xs flex-1 min-w-0"
                      style={inputStyle()}
                      placeholder="https://…"
                      value={neuerLink[m.id]?.url || ""}
                      onChange={(e) => setNeuerLink({ ...neuerLink, [m.id]: { ...neuerLink[m.id], url: e.target.value } })}
                      onKeyDown={(e) => e.key === "Enter" && linkHinzufuegen(m.id)}
                    />
                    <button onClick={() => linkHinzufuegen(m.id)} className="text-xs px-2 rounded border shrink-0" style={{ borderColor: T.accent, color: T.accent }}>
                      + Link
                    </button>
                  </div>
                </div>
              </div>
              ))
            )}
          </div>

          <div className="border rounded p-4 max-w-2xl bg-white" style={{ borderColor: T.line }}>
            <div className="text-xs font-medium mb-3" style={{ color: T.muted }}>Neue Methode anlegen</div>
            <div className="space-y-3">
              <Field label="Name">
                <input className="border rounded px-2 py-1 text-sm w-full" style={inputStyle()} value={neuMethode.name} onChange={(e) => setNeuMethode({ ...neuMethode, name: e.target.value })} />
              </Field>
              <Field label="Beschreibung">
                <div className="flex items-start gap-2">
                  <div className="flex-1">
                    <Beschreibungsfeld value={neuMethode.beschreibung} />
                  </div>
                  <button
                    onClick={() => setEditModal({ neu: true })}
                    title="Beschreibung bearbeiten"
                    className="text-xs shrink-0"
                    style={{ color: T.accent }}
                  >
                    ✎ Bearbeiten
                  </button>
                </div>
              </Field>
              <div>
                <div className="text-xs font-medium mb-1" style={{ color: T.muted }}>Fächer (Reihenfolge = Priorität für Auto-Vorschlag)</div>
                <div className="flex flex-wrap gap-1">
                  {faecher.map((f) => (
                    <button key={f.id} onClick={() => setNeuMethode({ ...neuMethode, faecherIds: toggleInList(neuMethode.faecherIds, f.id) })} className="text-xs px-2 py-1 rounded border" style={neuMethode.faecherIds.includes(f.id) ? { background: T.accentSoft, borderColor: T.accent } : { borderColor: T.line }}>
                      {f.name}{neuMethode.faecherIds.includes(f.id) && ` (${neuMethode.faecherIds.indexOf(f.id) + 1})`}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                <div className="text-xs font-medium mb-1" style={{ color: T.muted }}>Jahrgangsstufen</div>
                <div className="flex flex-wrap gap-1">
                  {JAHRGAENGE.map((jg) => (
                    <button key={jg} onClick={() => setNeuMethode({ ...neuMethode, jahrgaenge: toggleInList(neuMethode.jahrgaenge, jg) })} className="text-xs px-2 py-1 rounded border" style={neuMethode.jahrgaenge.includes(jg) ? { background: T.accentSoft, borderColor: T.accent } : { borderColor: T.line }}>
                      {jg}
                    </button>
                  ))}
                </div>
              </div>
              <Field label="Halbjahr (Empfehlung)">
                <div className="flex gap-2">
                  {[1, 2].map((h) => (
                    <button key={h} onClick={() => setNeuMethode({ ...neuMethode, halbjahr: h })} className="text-sm px-3 py-1 rounded border" style={neuMethode.halbjahr === h ? { background: T.ink, color: T.paper, borderColor: T.ink } : { borderColor: T.line }}>
                      {h}. Halbjahr
                    </button>
                  ))}
                </div>
              </Field>
              <Button onClick={addMethode} tone="accent">+ Methode anlegen</Button>
            </div>
          </div>
        </div>
      )}

      {stammTab === "import" && (
        <UntisImportView
          faecher={faecher} setFaecher={setFaecher}
          lehrer={lehrer} setLehrer={setLehrer}
          klassen={klassen} setKlassen={setKlassen}
          lerngruppen={lerngruppen} setLerngruppen={setLerngruppen}
          methoden={methoden} setMethoden={setMethoden}
          planungen={planungen} setPlanungen={setPlanungen}
        />
      )}

      {editModal && (
        <BeschreibungBearbeitenModal
          title={editModal.neu ? "Neue Methode – Beschreibung" : methoden.find((m) => m.id === editModal.methodeId)?.name || "Beschreibung"}
          value={editModal.neu ? neuMethode.beschreibung : methoden.find((m) => m.id === editModal.methodeId)?.beschreibung}
          onChange={(html) =>
            editModal.neu ? setNeuMethode({ ...neuMethode, beschreibung: html }) : updateMethode(editModal.methodeId, { beschreibung: html })
          }
          onClose={() => setEditModal(null)}
        />
      )}
      {pdfVorschau && <PdfVorschauModal material={pdfVorschau} onClose={() => setPdfVorschau(null)} />}
    </div>
  );
}

// ---------- Eine editierbare Zeile in der Lerngruppen-Übersicht ----------
function LerngruppeZeile({ g, faecher, lehrer, klassen, updateLerngruppe, removeLerngruppe }) {
  const [bestaetigen, setBestaetigen] = useState(false);
  const [wirdGeloescht, setWirdGeloescht] = useState(false);
  const fachObj = faecher.find((f) => f.id === g.fachId);

  const jaLoeschen = () => {
    // Zustand sofort anzeigen, die eigentliche Löschung (inkl. Neu-Rendern der ganzen Liste)
    // erst im nächsten Tick starten, damit der Browser "Wird gelöscht…" noch anzeigen kann,
    // bevor der spürbar dauernde Teil beginnt.
    setWirdGeloescht(true);
    setTimeout(() => removeLerngruppe(g.id), 20);
  };

  return (
    <div
      style={{ display: "grid", gridTemplateColumns: "110px 1fr 90px 90px" }}
      className="items-center px-3 py-2 text-sm gap-1"
    >
      <span className="truncate mc-mono text-xs" style={{ color: T.muted }}>
        {g.bezeichnung}
      </span>
      <span className="truncate">{fachObj?.name}</span>
      <select
        className="border rounded px-1.5 py-1 text-xs"
        style={inputStyle()}
        value={g.lehrerId}
        onChange={(e) => updateLerngruppe(g.id, { lehrerId: e.target.value })}
        disabled={wirdGeloescht}
      >
        {lehrer.map((l) => (
          <option key={l.id} value={l.id}>
            {l.kuerzel}
          </option>
        ))}
      </select>
      {wirdGeloescht ? (
        <span className="flex items-center gap-1.5 justify-self-end text-xs py-2" style={{ color: T.muted }}>
          <span
            className="inline-block w-3 h-3 rounded-full border-2 animate-spin"
            style={{ borderColor: T.line, borderTopColor: T.accent }}
          />
          Wird gelöscht…
        </span>
      ) : bestaetigen ? (
        <span className="flex items-center gap-1 justify-self-end">
          <button onClick={jaLoeschen} className="text-xs font-medium px-2 py-2.5" style={{ color: T.danger }}>
            Ja, löschen
          </button>
          <button onClick={() => setBestaetigen(false)} className="text-xs px-2 py-2.5" style={{ color: T.muted }}>
            Abbr.
          </button>
        </span>
      ) : (
        <button onClick={() => setBestaetigen(true)} className="text-xs justify-self-end px-2 py-2.5 -my-2.5" style={{ color: T.danger }}>
          Löschen
        </button>
      )}
    </div>
  );
}

// ---------- Untis-Import (Fächer, Lehrer, Klassen aus CSV-Export) ----------
function UntisImportView({ faecher, setFaecher, lehrer, setLehrer, klassen, setKlassen, lerngruppen, setLerngruppen, methoden, setMethoden, planungen, setPlanungen }) {
  const SEK1_MUSTER = /^(0[5-9]|10)[a-c]$/;

  const [gesamtZeilen, setGesamtZeilen] = useState(null);
  const [sek1Zeilen, setSek1Zeilen] = useState(null);
  const [fachZuordnung, setFachZuordnung] = useState({});
  const [modus, setModus] = useState("merge"); // 'merge' = nur Änderungen übernehmen, 'ueberschreiben' = vollständig ersetzen
  const [fehler, setFehler] = useState("");
  const [ergebnis, setErgebnis] = useState(null);

  const spalte = (row, ...namen) => {
    for (const n of namen) {
      if (row[n] !== undefined) return (row[n] || "").trim();
    }
    return "";
  };

  const onDatei = (e) => {
    const datei = e.target.files[0];
    if (!datei) return;
    setFehler("");
    setErgebnis(null);
    setSek1Zeilen(null);
    const reader = new FileReader();
    reader.onload = () => {
      Papa.parse(reader.result, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => {
          const zeilen = res.data
            .map((r) => ({
              subject: spalte(r, "subject", "Fach", "Subject"),
              teacher: spalte(r, "teacher", "Lehrer", "Teacher"),
              klassen: spalte(r, "klassen", "Klasse", "class", "Classes"),
            }))
            .filter((r) => r.subject && r.teacher);
          const gefiltert = zeilen.filter(
            (r) => r.klassen && r.klassen.split("~").every((k) => SEK1_MUSTER.test(k.trim()))
          );
          if (gefiltert.length === 0) {
            setFehler("Keine passenden Sek-I-Zeilen gefunden. Bitte Spaltenüberschriften der Datei prüfen (erwartet: subject, teacher, klassen).");
            return;
          }
          setGesamtZeilen(zeilen.length);
          setSek1Zeilen(gefiltert);
          const codes = [...new Set(gefiltert.map((r) => r.subject))].sort();
          const vorbelegung = {};
          codes.forEach((code) => {
            const bekannt = faecher.find((f) => f.kuerzel === code);
            vorbelegung[code] = { name: bekannt ? bekannt.name : code, skip: false, anzahl: gefiltert.filter((r) => r.subject === code).length };
          });
          setFachZuordnung(vorbelegung);
        },
      });
    };
    reader.readAsText(datei, "UTF-8");
  };

  const starteImport = () => {
    const neueFaecher = [...faecher];
    const neueLehrer = [...lehrer];
    let neueLerngruppen = [...lerngruppen];
    let angelegtFaecher = 0, umbenannteFaecher = 0, angelegtLehrer = 0;
    let angelegtLerngruppen = 0, aktualisierteLerngruppen = 0, entfernteLerngruppen = 0, uebersprungen = 0;
    let entfernteFaecher = 0, entfernteLehrer = 0;

    // Fach anlegen ODER (in jedem Modus) den in der Vorschau angepassten Namen übernehmen
    const findeOderErstelleFach = (code) => {
      let f = neueFaecher.find((x) => x.kuerzel === code);
      const gewuenschterName = fachZuordnung[code]?.name || code;
      if (!f) {
        f = { id: uid("f"), name: gewuenschterName, kuerzel: code, quelle: "untis" };
        neueFaecher.push(f);
        angelegtFaecher++;
      } else if (f.name !== gewuenschterName) {
        f.name = gewuenschterName;
        umbenannteFaecher++;
      }
      return f;
    };
    const findeOderErstelleLehrer = (kuerzel) => {
      let l = neueLehrer.find((x) => x.kuerzel === kuerzel);
      if (!l) {
        l = { id: uid("l"), name: "", kuerzel, email: "", quelle: "untis" };
        neueLehrer.push(l);
        angelegtLehrer++;
      }
      return l;
    };
    const findeKlasse = (token) => {
      const jahrgang = Number(token.slice(0, 2));
      const buchstabe = token.slice(2);
      return klassen.find((k) => k.jahrgang === jahrgang && k.buchstabe === buchstabe);
    };
    const schluessel = (fachKuerzel, klassenIds) => fachKuerzel + "|" + [...klassenIds].sort().join(",");

    // Zielzustand aus der Datei: ein Eintrag je Fach+Klassen-Kombination (unabhängig vom Lehrer,
    // damit ein Lehrerwechsel als Aktualisierung statt als Dublette erkannt wird)
    const ziele = new Map();
    sek1Zeilen.forEach((r) => {
      if (fachZuordnung[r.subject]?.skip) {
        uebersprungen++;
        return;
      }
      const klassenTokens = r.klassen.split("~").map((t) => t.trim());
      const klassenObjekte = klassenTokens.map(findeKlasse).filter(Boolean);
      if (klassenObjekte.length === 0) {
        uebersprungen++;
        return;
      }
      const klassenIds = klassenObjekte.map((k) => k.id);
      ziele.set(schluessel(r.subject, klassenIds), {
        fachCode: r.subject,
        klassenIds,
        klassenTokens,
        jahrgang: klassenObjekte[0].jahrgang,
        lehrerKuerzel: r.teacher,
      });
    });

    // Fächer für alle nicht ausgeschlossenen Codes anlegen bzw. umbenennen (vor dem Lerngruppen-Abgleich,
    // damit die folgenden Lookups die Fächer bereits vorfinden)
    [...new Set(sek1Zeilen.map((r) => r.subject))]
      .filter((code) => !fachZuordnung[code]?.skip)
      .forEach((code) => findeOderErstelleFach(code));

    // Bestehende, zuvor per Untis-Import angelegte Lerngruppen abgleichen: Lehrer aktualisieren,
    // falls geändert. Manuell angelegte Lerngruppen (ohne quelle:"untis") bleiben unangetastet.
    // Bewusst nicht auf Fächer dieser Datei eingeschränkt: sonst würden Lerngruppen zu Fächern, die
    // in dieser Datei komplett fehlen, im Überschreiben-Modus fälschlich nie entfernt.
    const gesehen = new Set();
    neueLerngruppen = neueLerngruppen.map((g) => {
      if (g.quelle !== "untis") return g;
      const fachKuerzel = neueFaecher.find((f) => f.id === g.fachId)?.kuerzel || "";
      const s = schluessel(fachKuerzel, g.klassenIds);
      const ziel = ziele.get(s);
      if (!ziel) return g;
      gesehen.add(s);
      const lehrerObj = findeOderErstelleLehrer(ziel.lehrerKuerzel);
      if (g.lehrerId !== lehrerObj.id) {
        aktualisierteLerngruppen++;
        return { ...g, lehrerId: lehrerObj.id, bezeichnung: ziel.klassenTokens.join("/") };
      }
      return g;
    });

    // Im Überschreiben-Modus: zuvor importierte Lerngruppen entfernen, die in dieser Datei nicht mehr auftauchen
    let entfernteLerngruppenIds = new Set();
    if (modus === "ueberschreiben") {
      const bleiben = [];
      neueLerngruppen.forEach((g) => {
        const fachKuerzel = neueFaecher.find((f) => f.id === g.fachId)?.kuerzel || "";
        const behalten = g.quelle !== "untis" || gesehen.has(schluessel(fachKuerzel, g.klassenIds));
        if (behalten) bleiben.push(g);
        else entfernteLerngruppenIds.add(g.id);
      });
      neueLerngruppen = bleiben;
      entfernteLerngruppen = entfernteLerngruppenIds.size;
    }

    // Neue Lerngruppen für alle Ziel-Kombinationen anlegen, die noch nicht abgedeckt sind
    ziele.forEach((ziel, s) => {
      if (gesehen.has(s)) return;
      const fachObj = neueFaecher.find((f) => f.kuerzel === ziel.fachCode);
      const lehrerObj = findeOderErstelleLehrer(ziel.lehrerKuerzel);
      neueLerngruppen.push({
        id: uid("lg"),
        fachId: fachObj.id,
        bezeichnung: ziel.klassenTokens.join("/"),
        jahrgang: ziel.jahrgang,
        lehrerId: lehrerObj.id,
        klassenIds: ziel.klassenIds,
        quelle: "untis",
      });
      angelegtLerngruppen++;
    });

    // Im Überschreiben-Modus: auch Fächer und Lehrkräfte vollständig abgleichen. Entfernt wird, was
    // zuvor selbst per Untis-Import angelegt wurde (quelle:"untis") und in dieser Datei nicht mehr
    // vorkommt – ohne Rücksicht darauf, ob noch etwas darauf verweist. Damit das keine kaputten
    // Verweise hinterlässt, werden abhängige Daten mit entfernt: Lerngruppen zu einem gelöschten Fach
    // sind zu diesem Zeitpunkt bereits oben herausgefiltert; zusätzlich werden gelöschte Fach-IDs aus
    // den Fächerlisten der Methoden entfernt und alle Planungen (Methoden-Zuordnungen) zu entfernten
    // Lerngruppen gelöscht. Nur händisch gepflegte Stammdaten (ohne quelle:"untis") bleiben unangetastet.
    let finaleFaecher = neueFaecher;
    let finaleLehrer = neueLehrer;
    let entfernteFachIds = new Set();
    if (modus === "ueberschreiben") {
      const zielFachCodes = new Set([...ziele.values()].map((z) => z.fachCode));
      const zielLehrerKuerzel = new Set([...ziele.values()].map((z) => z.lehrerKuerzel));

      finaleFaecher = neueFaecher.filter((f) => {
        if (f.quelle !== "untis" || zielFachCodes.has(f.kuerzel)) return true;
        entfernteFachIds.add(f.id);
        entfernteFaecher++;
        return false;
      });

      finaleLehrer = neueLehrer.filter((l) => {
        if (l.quelle !== "untis" || zielLehrerKuerzel.has(l.kuerzel)) return true;
        entfernteLehrer++;
        return false;
      });
    }

    // Kaskade: Methoden verlieren Verweise auf gelöschte Fächer, Planungen zu gelöschten
    // Lerngruppen werden komplett entfernt (inkl. bereits dokumentierter Durchführungen).
    let entfernteZuordnungen = 0;
    if (entfernteFachIds.size > 0) {
      setMethoden((prev) => prev.map((m) => ({ ...m, faecherIds: m.faecherIds.filter((id) => !entfernteFachIds.has(id)) })));
    }
    if (entfernteLerngruppenIds.size > 0) {
      const neuePlanungen = planungen.filter((p) => !entfernteLerngruppenIds.has(p.lerngruppeId));
      entfernteZuordnungen = planungen.length - neuePlanungen.length;
      setPlanungen(neuePlanungen);
    }

    setFaecher(finaleFaecher);
    setLehrer(finaleLehrer);
    setLerngruppen(neueLerngruppen);
    setErgebnis({
      angelegtFaecher, umbenannteFaecher, angelegtLehrer, entfernteFaecher, entfernteLehrer,
      angelegtLerngruppen, aktualisierteLerngruppen, entfernteLerngruppen, entfernteZuordnungen,
      uebersprungen, gesamt: sek1Zeilen.length,
    });
  };

  return (
    <div className="max-w-2xl">
      <div className="rounded-xl border p-4 bg-white mb-4" style={{ borderColor: T.line }}>
        <div className="mc-display text-sm font-semibold mb-1">CSV-Datei aus WebUntis wählen</div>
        <p className="text-xs mb-3" style={{ color: T.muted }}>
          Export "Unterricht" aus WebUntis (Administration → Export). Zeilen ohne Klasse (z.\u00a0B. Bereitschaft) und
          Klassen außerhalb der Sekundarstufe I (z.\u00a0B. EF, Q1, Q2, AG) werden automatisch ausgeschlossen.
        </p>
        <input type="file" accept=".csv,.txt" onChange={onDatei} className="text-sm" />
        {fehler && (
          <p className="text-xs mt-2" style={{ color: T.danger }}>
            {fehler}
          </p>
        )}
      </div>

      {sek1Zeilen && (
        <div className="rounded-xl border p-4 bg-white mb-4" style={{ borderColor: T.line }}>
          <div className="mc-display text-sm font-semibold mb-1">
            {sek1Zeilen.length} von {gesamtZeilen} Zeilen betreffen die Sekundarstufe I
          </div>
          <p className="text-xs mb-3" style={{ color: T.muted }}>
            Für jedes gefundene Fach-Kürzel: Anzeigename bestätigen/anpassen, oder Kürzel vom Import ausschließen
            (z.\u00a0B. für Differenzierungs- oder Wahlpflichtkürzel, die ihr nicht als eigenes Fach führen wollt).
          </p>
          <div className="rounded border divide-y" style={{ borderColor: T.line }}>
            {Object.keys(fachZuordnung).map((code) => (
              <div key={code} className="flex items-center gap-2 px-3 py-1.5 text-sm">
                <span className="mc-mono text-xs w-20 shrink-0" style={{ color: T.muted }}>
                  {code}
                </span>
                <input
                  className="border rounded px-2 py-1 text-sm flex-1"
                  style={inputStyle()}
                  value={fachZuordnung[code].name}
                  disabled={fachZuordnung[code].skip}
                  onChange={(e) =>
                    setFachZuordnung({ ...fachZuordnung, [code]: { ...fachZuordnung[code], name: e.target.value } })
                  }
                />
                <span className="text-xs w-20 shrink-0" style={{ color: T.muted }}>
                  {fachZuordnung[code].anzahl}× 
                </span>
                <label className="text-xs flex items-center gap-1 shrink-0" style={{ color: T.muted }}>
                  <input
                    type="checkbox"
                    checked={fachZuordnung[code].skip}
                    onChange={(e) =>
                      setFachZuordnung({ ...fachZuordnung, [code]: { ...fachZuordnung[code], skip: e.target.checked } })
                    }
                  />
                  ausschließen
                </label>
              </div>
            ))}
          </div>
          <div className="mt-4 mb-3">
            <div className="text-xs font-medium mb-1.5" style={{ color: T.muted }}>
              Modus
            </div>
            <div className="flex flex-col gap-2">
              <label className="flex items-start gap-2 text-sm cursor-pointer">
                <input
                  type="radio"
                  name="importmodus"
                  className="mt-0.5"
                  checked={modus === "merge"}
                  onChange={() => setModus("merge")}
                />
                <span>
                  <span className="font-medium">Nur Änderungen übernehmen</span>
                  <span className="block text-xs" style={{ color: T.muted }}>
                    Neue und geänderte Lerngruppen (z.\u00a0B. anderer Lehrer) werden übernommen. Bestehende
                    Lerngruppen bleiben erhalten, auch wenn sie in dieser Datei fehlen.
                  </span>
                </span>
              </label>
              <label className="flex items-start gap-2 text-sm cursor-pointer">
                <input
                  type="radio"
                  name="importmodus"
                  className="mt-0.5"
                  checked={modus === "ueberschreiben"}
                  onChange={() => setModus("ueberschreiben")}
                />
                <span>
                  <span className="font-medium">Bestehende Untis-Daten vollständig ersetzen</span>
                  <span className="block text-xs" style={{ color: T.muted }}>
                    Zuvor aus Untis importierte Lerngruppen, Fächer und Lehrkräfte, die in dieser Datei nicht mehr
                    auftauchen, werden entfernt – auch wenn ihnen bereits Methoden zugeordnet waren. Manuell
                    angelegte Stammdaten bleiben in jedem Fall unangetastet.
                  </span>
                </span>
              </label>
            </div>
          </div>
          <Button onClick={starteImport} tone="accent">
            Import durchführen
          </Button>
        </div>
      )}

      {ergebnis && (
        <div className="rounded-xl border p-4" style={{ borderColor: T.success, background: T.successSoft }}>
          <div className="mc-display text-sm font-semibold mb-1" style={{ color: T.success }}>
            Import abgeschlossen
          </div>
          <p className="text-xs" style={{ color: "#2E5A48" }}>
            {ergebnis.angelegtFaecher} neue Fächer{ergebnis.umbenannteFaecher > 0 && `, ${ergebnis.umbenannteFaecher} umbenannt`},{" "}
            {ergebnis.angelegtLehrer} neue Lehrkräfte (nur Kürzel – Name/E-Mail bitte in der Lehrer-Verwaltung ergänzen),{" "}
            {ergebnis.angelegtLerngruppen} neue{ergebnis.aktualisierteLerngruppen > 0 && ` und ${ergebnis.aktualisierteLerngruppen} aktualisierte`} Lerngruppen.
            {ergebnis.entfernteLerngruppen > 0 && ` ${ergebnis.entfernteLerngruppen} nicht mehr vorhandene Lerngruppen wurden entfernt.`}
            {(ergebnis.entfernteFaecher > 0 || ergebnis.entfernteLehrer > 0) &&
              ` Außerdem entfernt: ${ergebnis.entfernteFaecher} Fächer, ${ergebnis.entfernteLehrer} Lehrkräfte.`}
            {ergebnis.entfernteZuordnungen > 0 && ` Dabei wurden auch ${ergebnis.entfernteZuordnungen} bestehende Methoden-Zuordnungen zu diesen Lerngruppen zurückgenommen (inkl. eventuell bereits erfasster Durchführungen).`}
            {ergebnis.uebersprungen > 0 && ` ${ergebnis.uebersprungen} von ${ergebnis.gesamt} Zeilen wurden übersprungen (ausgeschlossenes Fach oder unbekannte Klasse).`}
          </p>
        </div>
      )}
    </div>
  );
}
