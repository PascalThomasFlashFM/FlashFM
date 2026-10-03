# -*- coding: utf-8 -*-
"""
Flash FM - Envoi des mails aux gagnants
Exécuter ce script sur votre ordinateur avec Python 3
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ============================================================
# CONFIGURATION - À modifier si besoin
# ============================================================
SMTP_SERVER = "ssl0.ovh.net"
SMTP_PORT   = 465
LOGIN       = "pascal@flashfm.fr"
PASSWORD    = "anton07"
FROM_ADDR   = "pascal@flashfm.fr"
FROM_NAME   = "Pascal Thomas Flash FM"

# Informations du jeu
NOM_JEU    = "Les Années 80, La Tournée"
DATE_HEURE = "samedi 6 juin 2026 à 20h00"
LIEU       = "Zénith Limoges Métropole"

# ============================================================
# LISTE DES GAGNANTS
# Format : (Prénom, Nom, email)
# ============================================================
GAGNANTS = [
    ("Jennifer",  "MALLE",      "missportugaise87@gmail.com"),
    ("Delphine",  "MODARD",     "delphinemodard01@gmail.com"),
    ("Julie",     "ARCIN",      "julie.arcin@yahoo.fr"),
    ("Catherine", "BOUDY",      "axelle.teddy@wanadoo.fr"),
    ("Sandrine",  "JAVERLIAC",  "sandrinemarcel972@gmail.com"),
]

# ============================================================
# FONCTIONS
# ============================================================

def construire_mail(prenom, nom, email_destinataire):
    plain = f"""Bonjour {prenom},

Suite à votre participation à notre jeu Flash FM « {NOM_JEU} », j'ai le plaisir de vous annoncer que vous avez été tiré au sort pour remporter 2 invitations !

Date : {DATE_HEURE}
Lieu : {LIEU}

Vous êtes inscrit sur une liste d'invités Flash FM. Vos places sont à retirer le soir du concert au guichet Invité sur la gauche du Zénith – Vous devrez vous munir d'une pièce d'identité.

Il n'est pas nécessaire de prendre la file d'attente du guichet principal, le guichet « invités » est sur la gauche à l'extérieur du Zénith.

Merci de me confirmer en retour la bonne réception de ce mail.

Vous retrouvez chaque semaine nos nouveaux jeux en écoutant Flash FM, sur notre site flashfm.fr, et sur notre application Flash FM France (application gratuite disponible sur tous les stores).

Merci de votre fidélité et bon spectacle !

PS : Si un jour vous êtes sollicité par téléphone par l'institut de Sondage Médiamétrie concernant les audiences radio, n'oubliez pas de répondre que vous écoutez Flash FM 😉

Vous souhaitez nous remercier pour ces places ou vous appréciez tout simplement Flash FM ?
Laissez-nous un avis sur notre fiche Google ici : https://g.page/r/CbZFpSFwfcq7EBM/review

Pascal Thomas
FLASH FM – Gérant SARL PROXIMEDIA – Directeur
30-32 Cours Gay Lussac – 87000 LIMOGES
05 55 31 00 00
pascal@flashfm.fr | http://www.flashfm.fr

FLASH FM 1ère radio musicale à LIMOGES et en HAUTE-VIENNE
Limoges : 89.9 – Saint-Junien : 98.4 – Brive : 99.9 – Uzerche : 99.9 – Guéret : 97.7 – Montmorillon : 95
DAB + : Haute-Vienne – Corrèze – Vienne – Et sur notre application Flash FM France
"""

    html = f"""<html><body style="font-family:Calibri,sans-serif;font-size:12pt;">
<p>Bonjour <strong>{prenom}</strong>,</p>

<p>Suite à votre participation à notre jeu Flash FM «&nbsp;<strong>{NOM_JEU}</strong>&nbsp;», j'ai le plaisir de vous annoncer que vous avez été tiré au sort pour remporter 2 invitations&nbsp;!</p>

<p><strong>Date&nbsp;: {DATE_HEURE}</strong><br>
<strong>Lieu&nbsp;: {LIEU}</strong></p>

<p><strong>Vous êtes inscrit sur une liste d'invités Flash FM.</strong> Vos places sont à retirer <strong>le soir du concert au guichet Invité sur la gauche du Zénith</strong> – Vous devrez vous munir d'une pièce d'identité.</p>

<p>Il n'est pas nécessaire de prendre la file d'attente du guichet principal, le guichet «&nbsp;invités&nbsp;» est sur la gauche à l'extérieur du Zénith.</p>

<br>
<p style="text-align:center;font-size:14pt;color:red;">
  <strong>Merci de me confirmer en retour la bonne réception de ce mail.</strong>
</p>
<br>

<p>Vous retrouvez chaque semaine nos nouveaux jeux en écoutant Flash FM, sur notre site
<a href="http://flashfm.fr">flashfm.fr</a>, et sur notre application Flash FM France
(application gratuite disponible sur tous les stores).</p>

<p>Merci de votre fidélité et bon spectacle&nbsp;!</p>
<br>

<p style="color:#7030A0;">
  <em><u>PS</u>&nbsp;: Si un jour vous êtes sollicité par téléphone par l'institut de Sondage
  Médiamétrie concernant les audiences radio, n'oubliez pas de répondre que vous écoutez
  Flash FM&nbsp;&#128521;</em>
</p>

<p style="color:#C00000;">
  <strong><em>Vous souhaitez nous remercier pour ces places ou vous appréciez tout simplement
  Flash FM&nbsp;?</em></strong>
  <em> <a href="https://g.page/r/CbZFpSFwfcq7EBM/review" style="color:#C00000;">
  Laissez-nous un avis sur notre fiche Google ici&nbsp;!</a></em>
</p>
<br>

<p>
  <strong><em><span style="font-size:14pt;color:#1F497D;">Pascal Thomas</span></em></strong><br>
  <strong><em><span style="font-size:14pt;color:#1F497D;">FLASH FM</span></em></strong><br>
  <em><span style="font-size:8pt;color:#7030A0;">Gérant SARL PROXIMEDIA</span></em><br>
  <em><span style="color:#17365D;">Directeur</span></em><br>
  <span style="color:#1F497D;">30-32 Cours Gay Lussac</span><br>
  <span style="color:#1F497D;">87000 LIMOGES</span><br>
  <span style="color:#4A442A;">05 55 31 00 00</span>
</p>

<p>
  <a href="mailto:pascal@flashfm.fr" style="color:purple;">pascal@flashfm.fr</a>&nbsp;|&nbsp;
  <a href="http://www.flashfm.fr" style="color:purple;">http://www.flashfm.fr</a>
</p>
<br>

<p>
  <em><span style="color:#1F497D;">FLASH FM 1<sup>ère</sup> radio musicale à LIMOGES et en HAUTE-VIENNE</span></em><br>
  Limoges&nbsp;: 89.9 – Saint-Junien&nbsp;: 98.4 – Brive&nbsp;: 99.9 – Uzerche&nbsp;: 99.9 – Guéret&nbsp;: 97.7 – Montmorillon&nbsp;: 95<br>
  DAB +&nbsp;: Haute-Vienne – Corrèze – Vienne<br>
  Et sur notre application Flash FM France
</p>

</body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Flash FM : {NOM_JEU}"
    msg["From"]    = f"{FROM_NAME} <{FROM_ADDR}>"
    msg["To"]      = email_destinataire
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html,  "html",  "utf-8"))
    return msg


def envoyer_tous():
    print(f"\nConnexion au serveur {SMTP_SERVER}:{SMTP_PORT} ...")
    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(LOGIN, PASSWORD)
            print("Connecté ✓\n")

            for prenom, nom, email in GAGNANTS:
                msg = construire_mail(prenom, nom, email)
                server.sendmail(FROM_ADDR, email, msg.as_string())
                print(f"  ✓ Mail envoyé à {prenom} {nom} <{email}>")

        print("\nTous les mails ont été envoyés avec succès !")

    except smtplib.SMTPAuthenticationError:
        print("ERREUR : Identifiant ou mot de passe incorrect.")
    except smtplib.SMTPException as e:
        print(f"ERREUR SMTP : {e}")
    except Exception as e:
        print(f"ERREUR : {e}")


# ============================================================
# PROGRAMME PRINCIPAL
# ============================================================
if __name__ == "__main__":
    print("=" * 55)
    print("  Flash FM – Envoi des mails gagnants")
    print(f"  Jeu : {NOM_JEU}")
    print(f"  Date : {DATE_HEURE}")
    print(f"  Lieu : {LIEU}")
    print("=" * 55)
    print(f"\n{len(GAGNANTS)} gagnants à contacter :")
    for prenom, nom, email in GAGNANTS:
        print(f"  - {prenom} {nom} <{email}>")

    print()
    reponse = input("Confirmer l'envoi ? (oui/non) : ").strip().lower()
    if reponse in ("oui", "o", "yes", "y"):
        envoyer_tous()
    else:
        print("Envoi annulé.")
