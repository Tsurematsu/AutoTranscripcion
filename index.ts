import ExcelJS, { Row } from "exceljs";
import path from "node:path";
import { fileURLToPath } from "url";

interface FilaTabla {
    Name: string;
    Text: string;
    Row: Row;
}

type ResultType = Array<Partial<Record<"mind" | "normal", string>>>;
type BoxTimeType = {
    boxSegment: FilaTabla[];
    textSegment: string[];
    Text: string;
    Type: "mind" | "normal";
}[];

interface FraseAsignada {
    bloque: FilaTabla | null;
    bloqueIndex: number | null;
    texto: string;
}

interface ItemScan {
    text: string;
    type: "mind" | "normal";
    frases: FraseAsignada[];
}

type Capa2Result = ItemScan[];

async function main() {
    console.clear();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.resolve(__dirname, "./docs/B24_All_Script_ES_final_EditK.xlsx"));
    const outputPath = path.resolve(__dirname, "./output", "CONSOLIDATED.xlsx");
    const sheet = workbook.getWorksheet(1);
    const newSheet = workbook.addWorksheet("Consolidated");
    if (!sheet) return;

    const column = sheet.getColumn("D");
    const cell = sheet.getCell('A1');
    const font = cell.font;

    const width = column.width || 8.43;
    const fontSize = font?.size || 12;
    const fontName = font?.name || "Arial";

    const approxCharsFit = () => {
        const fontFactors: any = {
            "calibri": 1.00,
            "arial": 1.02,
            "helvetica": 1.02,
            "verdana": 1.10,
            "tahoma": 1.07,
            "times new roman": 0.95,
            "courier new": 0.90,
        };
        const key = fontName.toLowerCase().trim();
        const factor = fontFactors[key] || 1.00;
        const chars = width * (11 / fontSize) * (1 / factor);
        return Math.floor(chars);
    };

    const maxChars = approxCharsFit();
    const maxLines = 2;

    let TempNameSet: string = "";
    const TempBlockText: FilaTabla[] = [];
    sheet.eachRow((row, rowNumber) => {
        const fila: FilaTabla = {
            Name: row.getCell("C").value?.toString().trim() || "",
            Text: row.getCell("D").value?.toString().trim()
                .replaceAll("\n", " ")
                .replaceAll(".¿", ". ¿")
                .replace(/\.([A-Za-z])/g, '. $1')
                .replace(/ +/g, ' ') || "",
            Row: row  // ✅ Agregado aquí
        };
        if (fila.Text.includes("[")) return;
        const strartTime = row.getCell("A").value?.toString().trim() || "";
        if (strartTime === "Start Time") return;
        if (fila.Name.length === 0) return;

        if (TempNameSet !== fila.Name) {
            if (TempNameSet !== "") {
                blockAnalysis(TempBlockText);
                TempBlockText.length = 0;
            }
            TempNameSet = fila.Name;
        }
        fila.Name = TempNameSet;
        TempBlockText.push(fila);
    });

    function blockAnalysis(block: FilaTabla[]) {
        const name = block[0].Name;
        const fullText = block.map(fila => fila.Text).join(" - ");

        console.log("_____________________________________________________");
        console.log("Name:" + name);

        console.log(fullText);
        const result1 = capa1(fullText);
        const result2 = capa2(result1, block);
        const result3 = capa3(result2, maxChars, maxLines);

    }

    function capa1(fullText: string) {
        const separeMind = fullText.replaceAll(" - ", " ").split("(");

        type BoxDialogType = Array<Partial<Record<"mind" | "normal", string>>>;
        const BoxDialog: BoxDialogType = [];

        if (separeMind.length > 0 && separeMind[0].trim().length > 0) {
            BoxDialog.push({ normal: separeMind[0].trim() });
        }

        for (let i = 1; i < separeMind.length; i++) {
            const element = separeMind[i];

            if (!element.includes(")")) {
                BoxDialog.push({ normal: element.trim() });
                continue;
            }

            const parts = element.trim().split(")");
            BoxDialog.push({ mind: parts[0].trim() });
            if (parts.length > 1 && parts[1].trim().length > 0) {
                BoxDialog.push({ normal: parts[1].trim() });
            }
        }

        const result: ResultType = [];
        const tempAccumulator: string[] = [];
        let flagKey: string = "";

        for (const element of BoxDialog) {
            const key = Object.keys(element)[0] as "mind" | "normal";
            const value = element[key] || "";

            if (flagKey !== key) {
                if (tempAccumulator.length > 0) {
                    result.push({ [flagKey]: tempAccumulator.join(" ") });
                    tempAccumulator.length = 0;
                }
                flagKey = key;
            }

            tempAccumulator.push(value);
        }

        if (tempAccumulator.length > 0) {
            result.push({ [flagKey]: tempAccumulator.join(" ") });
        }

        return result;
    }

    function capa2(boxResult: ResultType, blocks: FilaTabla[]): Capa2Result {
        const blokFrase: Capa2Result = [];

        for (const element of boxResult) {
            const text = Object.values(element)[0] || "";
            const itemScan: ItemScan = {
                text: text,
                type: Object.keys(element)[0] as "mind" | "normal",
                frases: []
            };

            const fraseAsignado: FraseAsignada[] = [];
            const partition = text.split('.');

            for (const frase of partition) {
                const itemFrase: FraseAsignada = {
                    bloque: null,
                    bloqueIndex: null,
                    texto: frase.trim()
                };

                if (itemFrase.texto.length === 0) continue;

                for (const block of blocks) {
                    const btxt = block.Text.replaceAll("(", "").replaceAll(")", "").replaceAll(" - ", " ");
                    if (itemFrase.bloque === null && btxt.includes(itemFrase.texto)) {
                        itemFrase.bloque = block;
                        itemFrase.bloqueIndex = blocks.indexOf(block);
                        break;
                    }
                    if (itemFrase.bloque === null && itemFrase.texto.includes(btxt)) {
                        itemFrase.bloque = block;
                        itemFrase.bloqueIndex = blocks.indexOf(block);
                        break;
                    }

                }
                fraseAsignado.push(itemFrase);
            }

            itemScan.frases = fraseAsignado;
            blokFrase.push(itemScan);
        }

        return blokFrase;
    }

    function capa3(blockText: Capa2Result, maxChars: number, maxLines: number) {
        // console.log(blockText);

        console.log(`\n------------- Max chars [${maxChars}] --------------\n`);


        for (const element of blockText) {

            const totalMaxChars = maxChars * maxLines
            const numRequireRows = Math.ceil(element.text.length / totalMaxChars)
            
            // const rangeCell = []
            // for (let index = 0; index < numRequireRows; index++) {
                
            // }

            for (const frase of element.frases) {
                if (frase.texto.length > maxChars) {
                    
                }
                console.log(`[${frase.texto.length <= maxChars}] ` + frase.texto);
            }
        }

    }

    // await workbook.xlsx.writeFile(outputPath);
}

main().catch(console.error);