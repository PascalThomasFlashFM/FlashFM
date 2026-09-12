import { AppConfig } from "../config";
import { bocirGraphQL } from "./client";

export interface FormField {
  id: string;
  label: string;
  type: string;
  position: number;
  mandatory: boolean;
}

interface FormQueryResult {
  Form: {
    id: string;
    title: string;
    fields: FormField[];
  };
}

const FORM_QUERY = `
  query($id: String!) { Form(id: $id) { id title fields { id label type position mandatory } } }
`;

// Labels exacts attendus sur le formulaire "La Voyance en direct".
// Le matching se fait par label, jamais par position ni par fieldId code en dur :
// les fieldId changent si le formulaire est recree.
export const EXPECTED_LABELS = {
  prenom: "Prénom",
  nom: "Nom",
  dateNaissance: "Date de naissance",
  telephone: "Téléphone",
  email: "Email",
  question: "La question que vous souhaitez poser à Marylou",
} as const;

export type FieldKey = keyof typeof EXPECTED_LABELS;

function normalizeLabel(label: string): string {
  return label.trim().toLowerCase();
}

export async function fetchFormFields(config: AppConfig): Promise<Map<FieldKey, FormField>> {
  const data = await bocirGraphQL<FormQueryResult>(config, FORM_QUERY, {
    id: config.bocir.formId,
  });

  const fields = data.Form.fields;
  const result = new Map<FieldKey, FormField>();

  for (const key of Object.keys(EXPECTED_LABELS) as FieldKey[]) {
    const expectedLabel = normalizeLabel(EXPECTED_LABELS[key]);
    const match = fields.find((f) => normalizeLabel(f.label) === expectedLabel);
    if (!match) {
      throw new Error(
        `Champ BOCIR introuvable pour le formulaire "${data.Form.title}" : ` +
          `label attendu "${EXPECTED_LABELS[key]}". Le formulaire a peut-etre ete modifie.`
      );
    }
    result.set(key, match);
  }

  return result;
}
