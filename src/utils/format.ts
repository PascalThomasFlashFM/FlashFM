const DDMMYY_FORMATTER = new Intl.DateTimeFormat("fr-FR", {
  day: "2-digit",
  month: "2-digit",
  year: "2-digit",
  timeZone: "Europe/Paris",
});

/** Timestamp Unix en millisecondes -> "JJ/MM/AA" */
export function msToDDMMYY(ms: number): string {
  return DDMMYY_FORMATTER.format(new Date(ms));
}

/** Date source API au format "AAAA-MM-JJ" -> "JJ/MM/AA" */
export function isoDateToDDMMYY(iso: string): string {
  const match = iso.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return iso; // format inattendu : on laisse la valeur brute plutot que de planter
  const [, year, month, day] = match;
  return `${day}/${month}/${year.slice(2)}`;
}

/** "0663598162" -> "06 63 59 81 62". Normalise un prefixe +33/0033 en 0. */
export function formatPhone(raw: string): string {
  let digits = raw.replace(/\D/g, "");
  if (digits.startsWith("33") && raw.trim().startsWith("+")) {
    digits = "0" + digits.slice(2);
  } else if (digits.startsWith("0033")) {
    digits = "0" + digits.slice(4);
  }
  const groups: string[] = [];
  for (let i = 0; i < digits.length; i += 2) {
    groups.push(digits.slice(i, i + 2));
  }
  return groups.join(" ");
}

/** Met la premiere lettre en majuscule, laisse le reste du prenom tel quel. */
export function capitalizeFirstLetter(value: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) return trimmed;
  return trimmed.charAt(0).toUpperCase() + trimmed.slice(1);
}
