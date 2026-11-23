const http = require("http");
const { URL } = require("url");

const DEFAULT_PORT = 3847;
const MAX_BODY_BYTES = 2 * 1024 * 1024;

/**
 * @typedef {{
 *   listProfiles: () => Promise<import("../lib/profiles.js").ProfileSummary[]>,
 *   findProfile: (profileId: string) => Promise<import("../lib/profiles.js").ProfileSummary | null>,
 *   findProfileByUsername: (username: string) => Promise<import("../lib/profiles.js").ProfileSummary | null>,
 *   findProfileCandidatesByUsername: (username: string) => Promise<import("../lib/profiles.js").ProfileSummary[]>,
 *   getJobContext: (profileId: string) => Promise<import("../lib/job-context-store.js").ProfileJobContext | null>,
 *   setJobContext: (profileId: string, patch: { jd?: string, company?: string }) => Promise<import("../lib/job-context-store.js").ProfileJobContext>,
 *   submitResumeBuild: (profile: import("../lib/profiles.js").ProfileSummary, input: { jd?: string, company?: string }) => Promise<{ profileId: string, username: string, jd: string, company: string, filePath: string, folderPath: string, jdPath: string, updatedAt: string }>,
 *   version?: string,
 *   port?: number,
 * }} ApiServerOptions
 */

/**
 * @param {http.ServerResponse} res
 * @param {number} status
 * @param {unknown} body
 */
function sendJson(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Length": Buffer.byteLength(payload),
  });
  res.end(payload);
}

/**
 * @param {import("http").IncomingMessage} req
 * @returns {Promise<Record<string, unknown>>}
 */
function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    /** @type {Buffer[]} */
    const chunks = [];
    let size = 0;

    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Request body too large."));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on("end", () => {
      if (chunks.length === 0) {
        resolve({});
        return;
      }
      try {
        const raw = Buffer.concat(chunks).toString("utf8").trim();
        resolve(raw ? JSON.parse(raw) : {});
      } catch (err) {
        reject(new Error("Invalid JSON body."));
      }
    });

    req.on("error", reject);
  });
}

/**
 * @param {string} pathname
 * @returns {{ resource: string, profileId?: string, action?: string }}
 */
function parseRoute(pathname) {
  const parts = pathname.replace(/\/+$/, "").split("/").filter(Boolean);
  if (parts.length === 0) {
    return { resource: "" };
  }
  if (parts[0] !== "api") {
    return { resource: parts.join("/") };
  }
  if (parts.length === 1) {
    return { resource: "api" };
  }
  if (parts[1] === "health") {
    return { resource: "health" };
  }
  if (parts[1] === "profiles" && parts.length === 2) {
    return { resource: "profiles" };
  }
  if (parts[1] === "job-context" && parts.length === 2) {
    return { resource: "job-context" };
  }
  if (parts[1] === "input" && parts.length === 2) {
    return { resource: "input" };
  }
  if (parts[1] === "profiles" && parts.length >= 3) {
    const profileId = decodeURIComponent(parts[2]);
    if (parts.length === 3) {
      return { resource: "profile", profileId };
    }
    if (parts[3] === "context") {
      return { resource: "profile-context", profileId };
    }
  }
  return { resource: parts.slice(1).join("/") };
}

/**
 * @param {ApiServerOptions} options
 */
function createApiServer(options) {
  const version = options.version || "0.0.0";
  const port =
    Number(process.env.RESUME_BUILDER_API_PORT || options.port || DEFAULT_PORT) ||
    DEFAULT_PORT;

  /** @type {http.Server} */
  const server = http.createServer(async (req, res) => {
    const method = String(req.method || "GET").toUpperCase();
    const url = new URL(req.url || "/", "http://127.0.0.1");
    const route = parseRoute(url.pathname);

    if (method === "OPTIONS") {
      sendJson(res, 204, {});
      return;
    }

    try {
      if (route.resource === "health" && method === "GET") {
        sendJson(res, 200, {
          ok: true,
          service: "resume-builder",
          version,
        });
        return;
      }

      if (route.resource === "profiles" && method === "GET") {
        const profiles = await options.listProfiles();
        sendJson(res, 200, { ok: true, profiles });
        return;
      }

      if (route.resource === "profile" && route.profileId && method === "GET") {
        const profile = await options.findProfile(route.profileId);
        if (!profile) {
          sendJson(res, 404, { ok: false, error: "Profile not found." });
          return;
        }
        sendJson(res, 200, { ok: true, profile });
        return;
      }

      if (
        route.resource === "profile-context" &&
        route.profileId &&
        method === "GET"
      ) {
        const profile = await options.findProfile(route.profileId);
        if (!profile) {
          sendJson(res, 404, { ok: false, error: "Profile not found." });
          return;
        }
        const context = await options.getJobContext(profile.id);
        sendJson(res, 200, {
          ok: true,
          profileId: profile.id,
          jd: context?.jd ?? "",
          company: context?.company ?? "",
          updatedAt: context?.updatedAt ?? null,
        });
        return;
      }

      if (
        route.resource === "profile-context" &&
        route.profileId &&
        method === "POST"
      ) {
        const profile = await options.findProfile(route.profileId);
        if (!profile) {
          sendJson(res, 404, { ok: false, error: "Profile not found." });
          return;
        }
        const body = await readJsonBody(req);
        if (
          !Object.prototype.hasOwnProperty.call(body, "jd") &&
          !Object.prototype.hasOwnProperty.call(body, "company")
        ) {
          sendJson(res, 400, {
            ok: false,
            error: "Provide at least one of: jd, company.",
          });
          return;
        }
        const context = await options.setJobContext(profile.id, {
          jd: Object.prototype.hasOwnProperty.call(body, "jd")
            ? String(body.jd ?? "")
            : undefined,
          company: Object.prototype.hasOwnProperty.call(body, "company")
            ? String(body.company ?? "")
            : undefined,
        });
        sendJson(res, 200, {
          ok: true,
          profileId: profile.id,
          jd: context.jd,
          company: context.company,
          updatedAt: context.updatedAt,
        });
        return;
      }

      if (route.resource === "job-context" && method === "POST") {
        const body = await readJsonBody(req);
        const profileId = String(body.profileId ?? body.profile ?? "").trim();
        if (!profileId) {
          sendJson(res, 400, { ok: false, error: "profileId is required." });
          return;
        }
        if (
          !Object.prototype.hasOwnProperty.call(body, "jd") &&
          !Object.prototype.hasOwnProperty.call(body, "company")
        ) {
          sendJson(res, 400, {
            ok: false,
            error: "Provide at least one of: jd, company.",
          });
          return;
        }
        const profile = await options.findProfile(profileId);
        if (!profile) {
          sendJson(res, 404, { ok: false, error: "Profile not found." });
          return;
        }
        const context = await options.setJobContext(profile.id, {
          jd: Object.prototype.hasOwnProperty.call(body, "jd")
            ? String(body.jd ?? "")
            : undefined,
          company: Object.prototype.hasOwnProperty.call(body, "company")
            ? String(body.company ?? "")
            : undefined,
        });
        sendJson(res, 200, {
          ok: true,
          profileId: profile.id,
          jd: context.jd,
          company: context.company,
          updatedAt: context.updatedAt,
        });
        return;
      }

      if (route.resource === "input" && method === "POST") {
        const body = await readJsonBody(req);
        const username = String(
          body.username ?? body.user ?? body.name ?? body.label ?? "",
        ).trim();
        const profileId = String(body.profileId ?? body.profile ?? "").trim();
        const jd = Object.prototype.hasOwnProperty.call(body, "jd")
          ? String(body.jd ?? "")
          : undefined;
        const company = Object.prototype.hasOwnProperty.call(body, "company")
          ? String(body.company ?? "")
          : undefined;

        if (!username && !profileId) {
          sendJson(res, 400, {
            ok: false,
            error: "username is required.",
          });
          return;
        }
        if (jd === undefined && company === undefined) {
          sendJson(res, 400, {
            ok: false,
            error: "Provide at least one of: jd, company.",
          });
          return;
        }

        let profile = null;
        if (profileId) {
          profile = await options.findProfile(profileId);
        } else {
          profile = await options.findProfileByUsername(username);
          if (!profile) {
            const candidates = await options.findProfileCandidatesByUsername(
              username,
            );
            if (candidates.length > 1) {
              sendJson(res, 409, {
                ok: false,
                error: "Multiple profiles match username.",
                username,
                profiles: candidates,
              });
              return;
            }
          }
        }

        if (!profile) {
          sendJson(res, 404, {
            ok: false,
            error: "Profile not found for username.",
            username: username || undefined,
            profileId: profileId || undefined,
          });
          return;
        }

        try {
          const result = await options.submitResumeBuild(profile, {
            jd: jd ?? "",
            company: company ?? "",
          });
          sendJson(res, 200, {
            ok: true,
            profileId: result.profileId,
            username: result.username,
            jd: result.jd,
            company: result.company,
            filePath: result.filePath,
            folderPath: result.folderPath,
            jdPath: result.jdPath || undefined,
            updatedAt: result.updatedAt,
          });
        } catch (err) {
          const message =
            err && typeof err === "object" && "message" in err
              ? String(err.message)
              : "Resume build failed.";
          const status = /timed out/i.test(message) ? 504 : 500;
          sendJson(res, status, {
            ok: false,
            error: message,
            profileId: profile.id,
            username: profile.label,
          });
        }
        return;
      }

      sendJson(res, 404, { ok: false, error: "Not found." });
    } catch (err) {
      const message =
        err && typeof err === "object" && "message" in err
          ? String(err.message)
          : "Internal server error.";
      const status = /invalid json|too large|required/i.test(message) ? 400 : 500;
      sendJson(res, status, { ok: false, error: message });
    }
  });

  return {
    server,
    port,
    start() {
      return new Promise((resolve, reject) => {
        server.once("error", reject);
        server.listen(port, "127.0.0.1", () => {
          server.removeListener("error", reject);
          console.log(`Resume Builder API listening on http://127.0.0.1:${port}`);
          resolve(port);
        });
      });
    },
    stop() {
      return new Promise((resolve, reject) => {
        server.close((err) => {
          if (err) {
            reject(err);
            return;
          }
          resolve();
        });
      });
    },
  };
}

module.exports = {
  DEFAULT_PORT,
  createApiServer,
};
