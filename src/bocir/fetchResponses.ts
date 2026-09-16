import { AppConfig } from "../config";
import { bocirGraphQL } from "./client";
import { FieldKey, FormField } from "./fetchFormFields";

interface RawResponseItem {
  fieldId: string;
  position: number;
  label: string;
  value: string;
}

interface RawFormResponse {
  id: string;
  formId: string;
  submissionDate: string; // timestamp Unix en ms, en string cote API
  comment: string | null;
  responses: RawResponseItem[];
}

interface ResponsesQueryResult {
  allFormResponses: RawFormResponse[];
  _allFormResponsesMeta: { count: number };
}

const RESPONSES_QUERY = `
  query($filter: FormResponseFilter, $perPage: Float, $page: Float) {
    allFormResponses(filter: $filter, perPage: $perPage, page: $page, sortField: "submissionDate", sortOrder: "ASC") {
      id
      formId
      submissionDate
      comment
      responses { fieldId position label value }
    }
    _allFormResponsesMeta(filter: $filter) { count }
  }
`;

const PER_PAGE = 50;

export interface NewRegistration {
  bocirId: string;
  submissionDate: number;
  prenom: string;
  nom: string;
  email: string;
  telephone: string;
  dateNaissance: string;
  question: string;
}

function normalizeLabel(label: string): string {
  return label.trim().toLowerCase();
}

function extractValue(
  item: RawFormResponse,
  fieldKey: FieldKey,
  fieldMap: Map<FieldKey, FormField>
): string {
  const field = fieldMap.get(fieldKey)!;
  const byFieldId = item.responses.find((r) => r.fieldId === field.id);
  if (byFieldId) return byFieldId.value ?? "";
  const byLabel = item.responses.find(
    (r) => normalizeLabel(r.label) === normalizeLabel(field.label)
  );
  return byLabel?.value ?? "";
}

/**
 * Recupere toutes les inscriptions dont submissionDate > lastProcessed, triees par
 * submissionDate ascendant. L'API BOCIR ne permet pas de filtrer par date : on doit
 * parcourir toutes les pages (pagination page=0..n, jamais page=1 en premiere page)
 * et ignorer localement celles deja traitees.
 */
export async function fetchNewResponses(
  config: AppConfig,
  fieldMap: Map<FieldKey, FormField>,
  lastProcessedSubmissionDate: number
): Promise<NewRegistration[]> {
  const results: NewRegistration[] = [];
  let page = 0;

  while (true) {
    const data = await bocirGraphQL<ResponsesQueryResult>(config, RESPONSES_QUERY, {
      filter: { formId: config.bocir.formId },
      perPage: PER_PAGE,
      page,
    });

    const batch = data.allFormResponses;
    if (!batch || batch.length === 0) break;

    for (const item of batch) {
      const submissionDate = parseInt(item.submissionDate, 10);
      if (submissionDate > lastProcessedSubmissionDate) {
        results.push({
          bocirId: item.id,
          submissionDate,
          prenom: extractValue(item, "prenom", fieldMap),
          nom: extractValue(item, "nom", fieldMap),
          email: extractValue(item, "email", fieldMap),
          telephone: extractValue(item, "telephone", fieldMap),
          dateNaissance: extractValue(item, "dateNaissance", fieldMap),
          question: extractValue(item, "question", fieldMap),
        });
      }
    }

    if (batch.length < PER_PAGE) break;
    page += 1;
  }

  results.sort((a, b) => a.submissionDate - b.submissionDate);
  return results;
}
