import { AppConfig } from "../config";

interface GraphQLResponse<T> {
  data?: T;
  errors?: Array<{ message: string }>;
}

export async function bocirGraphQL<T>(
  config: AppConfig,
  query: string,
  variables: Record<string, unknown>
): Promise<T> {
  const res = await fetch(config.bocir.endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": config.bocir.apiKey,
      authorization: `Bearer ${config.bocir.bearerToken}`,
    },
    body: JSON.stringify({ query, variables }),
  });

  if (!res.ok) {
    throw new Error(`BOCIR HTTP ${res.status} ${res.statusText}`);
  }

  const json = (await res.json()) as GraphQLResponse<T>;

  if (json.errors && json.errors.length > 0) {
    const messages = json.errors.map((e) => e.message).join("; ");
    throw new Error(
      `BOCIR GraphQL error: ${messages} ` +
        `(une erreur "Unexpected error" generique signale souvent un champ mal nomme dans la requete, pas un probleme d'auth)`
    );
  }

  if (!json.data) {
    throw new Error("BOCIR: reponse sans data et sans erreur explicite");
  }

  return json.data;
}
