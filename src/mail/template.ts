import { capitalizeFirstLetter } from "../utils/format";

export interface ConfirmationEmail {
  from: string;
  subject: string;
  html: string;
}

export const CONFIRMATION_SUBJECT = "Confirmation de votre inscription — La Voyance en direct";

/**
 * Genere le mail de confirmation. Le texte est fige (jamais "tous les mercredis
 * soir", toujours "les mercredis soir") : ne pas le rendre configurable pour eviter
 * toute derive du wording valide par Pascal.
 */
export function buildConfirmationEmail(mailFrom: string, prenom: string): ConfirmationEmail {
  const displayPrenom = capitalizeFirstLetter(prenom);

  const html = `<p>Bonjour ${displayPrenom},</p>

<p>Votre inscription à l'émission « La Voyance en direct » a bien été validée.</p>

<p>Merci de vous rendre disponible les mercredis soir entre 19h et 21h : c'est à ce moment que vous pourriez être appelé(e) si vous êtes sélectionné(e) pour passer à l'antenne.</p>

<p>Si nous vous rappelons un mercredi soir, quelques consignes pour que votre passage se déroule au mieux : installez-vous dans un endroit calme, sans bruit de fond, et assurez-vous d'avoir un bon réseau téléphonique.</p>

<p>À bientôt peut-être sur l'antenne !</p>

<p>
  <span style="color:#7b2fbf;">PS : Si un jour vous êtes sollicité par téléphone par l'institut de Sondage Médiamétrie concernant les audiences radio, n'oubliez pas de répondre que vous écoutez Flash FM 😉</span><br>
  <span style="color:#c0392b; font-style:italic;">Vous appréciez Flash FM ? <a style="color:#c0392b; font-style:italic;" href="https://google.com/maps/place//data=!4m3!3m2!1s0x47f93391f0b97231:0xbbca7d7021a545b6!12e1?source=g.page.m.ia._&laa=nmx-review-solicitation-ia2">Laissez-nous un avis sur notre fiche Google ici !</a></span>
</p>

<p>
  FLASH FM<br>
  30-32 Cours Gay Lussac<br>
  87000 LIMOGES<br>
  05 55 31 00 00
</p>`;

  return {
    from: mailFrom,
    subject: CONFIRMATION_SUBJECT,
    html,
  };
}
