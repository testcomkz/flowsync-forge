import { toDateInputValue } from "@/components/ui/date-input";

export type StageKey = "rattling" | "external" | "hydro" | "mpi" | "drift" | "emi" | "marking";
export type ScrapKey = "rattling" | "external" | "jetting" | "mpi" | "drift" | "emi";

export interface TubingRecord {
  id: string;
  client: string;
  wo_no: string;
  batch: string;
  status: string;
  diameter: string;
  qty: string;
  pipe_from: string;
  pipe_to: string;
  rack: string;
  arrival_date: string;
  class_1: string;
  class_2: string;
  class_3: string;
  repair: string;
  scrapTotal: string;
  start_date: string;
  end_date: string;
  load_out_date: string;
  act_no_oper: string;
  act_date: string;
  quantities: Partial<Record<StageKey, string>>;
  scrap: Partial<Record<ScrapKey, string>>;
  originalClient: string;
  originalWo: string;
  originalBatch: string;
}

const normalize = (value: unknown) => (value === null || value === undefined ? "" : String(value).trim());
const normalizeLower = (value: unknown) => normalize(value).toLowerCase();
const canonicalize = (header: string) => header.replace(/[\s-]+/g, "_").replace(/_{2,}/g, "_");

export const sanitizeNumberString = (value: string) => value.replace(/[^0-9-]/g, "");

export const computePipeTo = (pipeFrom: string, qty: string) => {
  const parsedFrom = Number.parseInt(sanitizeNumberString(pipeFrom), 10);
  const parsedQty = Number.parseInt(sanitizeNumberString(qty), 10);
  if (Number.isNaN(parsedFrom) || Number.isNaN(parsedQty)) {
    return "";
  }
  const pipeTo = parsedFrom + parsedQty - 1;
  return pipeTo.toString();
};

export const parseTubingRecords = (data: any[]): TubingRecord[] => {
  if (!Array.isArray(data) || data.length < 2) {
    return [];
  }

  const headers = data[0] as unknown[];

  const findIndex = (matcher: (normalized: string, canonical: string) => boolean) =>
    headers.findIndex(header => {
      const normalized = normalizeLower(header);
      const canonical = canonicalize(normalized);
      return matcher(normalized, canonical);
    });

  const clientIndex = findIndex(header => header.includes("client"));
  const woIndex = findIndex(header => header.includes("wo"));
  const batchIndex = findIndex(header => header.includes("batch"));
  const statusIndex = findIndex(header => header.includes("status"));
  const diameterIndex = findIndex(header => header.includes("diameter") || header.includes("диаметр"));
  const qtyIndex = findIndex((header, canonical) =>
    canonical === "qty" ||
    canonical === "quantity" ||
    (canonical.includes("qty") &&
      !canonical.includes("scrap") &&
      !canonical.includes("rattling") &&
      !canonical.includes("external") &&
      !canonical.includes("hydro") &&
      !canonical.includes("mpi") &&
      !canonical.includes("drift") &&
      !canonical.includes("emi") &&
      !canonical.includes("marking"))
  );
  const pipeFromIndex = findIndex((header, canonical) => canonical.includes("pipe_from"));
  const pipeToIndex = findIndex((header, canonical) => canonical.includes("pipe_to"));
  const rackIndex = findIndex((header, canonical) => canonical.includes("rack"));
  const arrivalDateIndex = findIndex((header, canonical) => canonical.includes("arrival_date"));
  const class1Index = findIndex((header, canonical) => canonical.includes("class_1") || header.includes("class 1"));
  const class2Index = findIndex((header, canonical) => canonical.includes("class_2") || header.includes("class 2"));
  const class3Index = findIndex((header, canonical) => canonical.includes("class_3") || header.includes("class 3"));
  const repairIndex = findIndex(header => header.includes("repair"));
  const scrapIndex = findIndex((header, canonical) => canonical === "scrap" || canonical.endsWith("_scrap"));
  const startDateIndex = findIndex((header, canonical) => canonical.includes("start_date"));
  const endDateIndex = findIndex((header, canonical) => canonical.includes("end_date"));
  const loadOutDateIndex = findIndex((header, canonical) => canonical.includes("load_out_date") || canonical.includes("loadoutdate"));
  const actNoOperIndex = findIndex((header, canonical) => canonical.includes("act_no_oper") || canonical.includes("actnooper"));
  const actDateIndex = findIndex((header, canonical) => canonical.includes("act_date") || canonical.includes("actdate"));

  const rattlingQtyIndex = findIndex((header, canonical) => canonical.includes("rattling_qty") && !canonical.includes("scrap"));
  const externalQtyIndex = findIndex((header, canonical) => canonical.includes("external_qty") && !canonical.includes("scrap"));
  const hydroQtyIndex = findIndex((header, canonical) =>
    (canonical.includes("hydro_qty") || canonical.includes("jetting_qty")) && !canonical.includes("scrap")
  );
  const mpiQtyIndex = findIndex((header, canonical) => canonical.includes("mpi_qty") && !canonical.includes("scrap"));
  const driftQtyIndex = findIndex((header, canonical) => canonical.includes("drift_qty") && !canonical.includes("scrap"));
  const emiQtyIndex = findIndex((header, canonical) => canonical.includes("emi_qty") && !canonical.includes("scrap"));
  const markingQtyIndex = findIndex((header, canonical) => canonical.includes("marking_qty"));

  const rattlingScrapIndex = findIndex((header, canonical) => canonical.includes("rattling_scrap"));
  const externalScrapIndex = findIndex((header, canonical) => canonical.includes("external_scrap"));
  const jettingScrapIndex = findIndex((header, canonical) => canonical.includes("jetting_scrap"));
  const mpiScrapIndex = findIndex((header, canonical) => canonical.includes("mpi_scrap"));
  const driftScrapIndex = findIndex((header, canonical) => canonical.includes("drift_scrap"));
  const emiScrapIndex = findIndex((header, canonical) => canonical.includes("emi_scrap"));

  return (data.slice(1) as unknown[][])
    .map((row, rowIndex) => {
      const client = normalize(clientIndex >= 0 ? row[clientIndex] : "");
      const wo_no = normalize(woIndex >= 0 ? row[woIndex] : "");
      const batch = normalize(batchIndex >= 0 ? row[batchIndex] : "");
      if (!client || !wo_no || !batch) {
        return null;
      }

      const status = normalize(statusIndex >= 0 ? row[statusIndex] : "");
      const diameter = normalize(diameterIndex >= 0 ? row[diameterIndex] : "");
      const qty = normalize(qtyIndex >= 0 ? row[qtyIndex] : "");
      const pipe_from = normalize(pipeFromIndex >= 0 ? row[pipeFromIndex] : "");
      const pipe_to = normalize(pipeToIndex >= 0 ? row[pipeToIndex] : "");
      const rack = normalize(rackIndex >= 0 ? row[rackIndex] : "");
      const arrival_date = toDateInputValue(arrivalDateIndex >= 0 ? row[arrivalDateIndex] : "");
      const class_1 = normalize(class1Index >= 0 ? row[class1Index] : "");
      const class_2 = normalize(class2Index >= 0 ? row[class2Index] : "");
      const class_3 = normalize(class3Index >= 0 ? row[class3Index] : "");
      const repair = normalize(repairIndex >= 0 ? row[repairIndex] : "");
      const scrapTotal = normalize(scrapIndex >= 0 ? row[scrapIndex] : "");
      const start_date = toDateInputValue(startDateIndex >= 0 ? row[startDateIndex] : "");
      const end_date = toDateInputValue(endDateIndex >= 0 ? row[endDateIndex] : "");
      const load_out_date = toDateInputValue(loadOutDateIndex >= 0 ? row[loadOutDateIndex] : "");
      const act_no_oper = normalize(actNoOperIndex >= 0 ? row[actNoOperIndex] : "");
      const act_date = toDateInputValue(actDateIndex >= 0 ? row[actDateIndex] : "");

      const quantities: Partial<Record<StageKey, string>> = {
        rattling: normalize(rattlingQtyIndex >= 0 ? row[rattlingQtyIndex] : ""),
        external: normalize(externalQtyIndex >= 0 ? row[externalQtyIndex] : ""),
        hydro: normalize(hydroQtyIndex >= 0 ? row[hydroQtyIndex] : ""),
        mpi: normalize(mpiQtyIndex >= 0 ? row[mpiQtyIndex] : ""),
        drift: normalize(driftQtyIndex >= 0 ? row[driftQtyIndex] : ""),
        emi: normalize(emiQtyIndex >= 0 ? row[emiQtyIndex] : ""),
        marking: normalize(markingQtyIndex >= 0 ? row[markingQtyIndex] : ""),
      };

      const scrap: Partial<Record<ScrapKey, string>> = {
        rattling: normalize(rattlingScrapIndex >= 0 ? row[rattlingScrapIndex] : ""),
        external: normalize(externalScrapIndex >= 0 ? row[externalScrapIndex] : ""),
        jetting: normalize(jettingScrapIndex >= 0 ? row[jettingScrapIndex] : ""),
        mpi: normalize(mpiScrapIndex >= 0 ? row[mpiScrapIndex] : ""),
        drift: normalize(driftScrapIndex >= 0 ? row[driftScrapIndex] : ""),
        emi: normalize(emiScrapIndex >= 0 ? row[emiScrapIndex] : ""),
      };

      return {
        id: `${rowIndex}-${client}-${wo_no}-${batch}`,
        client,
        wo_no,
        batch,
        status,
        diameter,
        qty,
        pipe_from,
        pipe_to,
        rack,
        arrival_date,
        class_1,
        class_2,
        class_3,
        repair,
        scrapTotal,
        start_date,
        end_date,
        load_out_date,
        act_no_oper,
        act_date,
        quantities,
        scrap,
        originalClient: client,
        originalWo: wo_no,
        originalBatch: batch,
      } satisfies TubingRecord;
    })
    .filter((value): value is TubingRecord => Boolean(value));
};
