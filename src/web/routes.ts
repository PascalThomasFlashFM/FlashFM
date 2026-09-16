import { Router } from "express";
import { buildPreview, excludeRegistration, sendAllPending, sendRegistrationMail } from "../actions";
import { AppConfig } from "../config";
import * as db from "../db";
import { logger } from "../logger";
import { runSync } from "../sync/runSync";

export function createApiRouter(config: AppConfig): Router {
  const router = Router();

  router.get("/registrations", (req, res) => {
    const status = req.query.status as db.MailStatus | undefined;
    const regs = db.listRegistrations(status);
    res.json(regs.map((r) => buildPreview(r, config.smtp.mailFrom)));
  });

  router.post("/registrations/:id/exclude", (req, res) => {
    try {
      const reason = typeof req.body?.reason === "string" ? req.body.reason : null;
      const updated = excludeRegistration(req.params.id, reason);
      res.json(updated);
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : String(err) });
    }
  });

  router.post("/registrations/:id/send", async (req, res) => {
    try {
      const result = await sendRegistrationMail(config, req.params.id);
      if (result.ok) {
        res.json({ ok: true });
      } else {
        res.status(502).json({ ok: false, error: result.error });
      }
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : String(err) });
    }
  });

  router.post("/send-all", async (req, res) => {
    const result = await sendAllPending(config);
    res.json(result);
  });

  router.get("/history", (req, res) => {
    res.json({
      envoyes: db.listRegistrations("envoye"),
      exclus: db.listRegistrations("exclu"),
      runs: db.listSyncRuns(30),
    });
  });

  router.post("/sync", async (req, res) => {
    try {
      const summary = await runSync(config);
      res.json(summary);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      logger.error(`Echec sync declenchee depuis l'UI : ${message}`);
      res.status(500).json({ error: message });
    }
  });

  return router;
}
