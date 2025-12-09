import ExcelJS from "exceljs";
import path from "node:path";
import { fileURLToPath } from "url";

interface FilaTabla {
    StartTime: string;
    EndTime: string;
    Name: string;
    Text: string;
    RevisedText: string;
    WPS: string;
    TotalStartTime: string;
    TotalEndTime: string;
}

type ResultType = Array<Partial<Record<"mind" | "normal", string>>>;
type BoxTimeType = {
    boxSegment: FilaTabla[];
    textSegment: string[];
    Text: string;
    Type: "mind" | "normal";
}[];

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

    let TempNameSet: string = "";
    const TempBlockText: FilaTabla[] = [];
    sheet.eachRow((row, rowNumber) => {
        const fila: FilaTabla = {
            StartTime: row.getCell("A").value?.toString().trim() || "",
            EndTime: row.getCell("B").value?.toString().trim() || "",
            Name: row.getCell("C").value?.toString().trim() || "",
            Text: row.getCell("D").value?.toString().trim()
                .replaceAll("\n", " ")
                .replaceAll(".¿", ". ¿")
                .replace(/\.([A-Za-z])/g, '. $1')
                .replace(/ +/g, ' ') || "",
            RevisedText: row.getCell("E").value?.toString().trim() || "",
            WPS: row.getCell("F").value?.toString().trim() || "",
            TotalStartTime: row.getCell("G").value?.toString().trim() || "",
            TotalEndTime: row.getCell("H").value?.toString().trim() || ""
        };
        if (fila.Text.includes("[")) return;
        if (fila.StartTime === "Start Time") return;
        if (fila.Name.length === 0) return;

        if (TempNameSet !== fila.Name) {
            if (TempNameSet !== "") {
                blockAnalysis(TempBlockText);
                TempBlockText.length = 0; // Clear the array for the next block
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
        // console.log(result1);

        const result2 = capa2(result1, block);
        // console.log(result2);

    }
    function capa1(fullText: string) {

        const separeMind = fullText.replaceAll(" - ", " ").split("(")
        if (separeMind.length > 1) separeMind.shift();

        type BoxDialogType = Array<Partial<Record<"mind" | "normal", string>>>;
        const BoxDialog: BoxDialogType = [];

        for (const element of separeMind) {
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
                // Si hay elementos acumulados, agrégalos al resultado
                if (tempAccumulator.length > 0) {
                    result.push({ [flagKey]: tempAccumulator.join(" ") });
                    tempAccumulator.length = 0;
                }
                flagKey = key;
            }

            tempAccumulator.push(value);
        }

        // ¡IMPORTANTE! Agregar el último grupo acumulado
        if (tempAccumulator.length > 0) result.push({ [flagKey]: tempAccumulator.join(" ") });

        return result;
    }
    function capa2(boxResult: ResultType, blocks: FilaTabla[]) {
        console.log(boxResult);

        console.log("-------------------");
        

        // const BoxTime: BoxTimeType = [];
        // for (const block of blocks) {
        //     const textBlock = block.Text.replaceAll("(", "").replaceAll(")", "").replace(/ +/g, ' ').trim();
        //     // console.log(textBlock);
        //     for (const frase of boxResult) {
        //         const key = Object.keys(frase)[0] as "mind" | "normal";
        //         const text = frase[key] || "";
        //         if (text.includes(textBlock) || textBlock.includes(text)) {
        //             console.log(`[${boxResult.indexOf(frase)}][${key}] => ${textBlock}`);
        //         }
        //     }
        // }
        
        // for (const frase of boxResult) {
        //     const key = Object.keys(frase)[0] as "mind" | "normal";
        //     const text = frase[key] || "";
        //     console.log(text);
        // }


        // console.log(boxResult);

    }

    // await workbook.xlsx.writeFile(outputPath);
}



main().catch(console.error);