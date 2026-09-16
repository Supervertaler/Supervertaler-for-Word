"""The local engine, first cut: serves the task pane and answers its questions
from a Supervertaler database, read-only.

    python engine/engine.py --db "D:\\path\\supervertaler.db" --termbases 4 [--port 3000]

Endpoints (all JSON):
    GET  /api/status                      what is loaded
    GET  /api/terms?text=&src=nl&tgt=en[&tbs=13,103]   termbase hits for a sentence
    GET  /api/tm?text=&src=nl&tgt=en[&tms=BEIJER,x]     TM candidates (the pane scores them)
    tbs / tms restrict the lookup to the resources the document has switched on.
    POST /api/terms  {src, tgt, src_lang, tgt_lang, termbase_id}
                                          add a term to a termbase in the database

Reads use a read-only connection. Adding a term is the one write, done with
the same columns the Trados plugin uses, so Studio sees the term at once.
--project-termbase names the default destination (id or name).

Scale: terms for the configured termbases are held in memory (51,000 for
BEIJER, trivial). TM candidates come from the database's own FTS5 index on
translation units, 30 ms on 877,000 units.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
import threading
import uuid
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(HERE, "..", "spike", "taskpane")
PENDING = os.path.join(HERE, "pending_terms.json")

_word = re.compile(r"[\w'’-]+", re.UNICODE)

# Function words that some termbases carry as entries ("voor -> for"). A term
# made only of these is noise in a lens, whatever the termbase says.
STOP = set("""de het een en van voor dit dat op in te met is zijn aan bij of om als
door uit naar over ook nog dan wordt worden werd hij zij ze het er
the a an and of for this that with to at by on in is are be as or from into""".split())
def trivial(key: str) -> bool:
    ws = key.split()
    return all(w in STOP for w in ws) or (len(ws) == 1 and len(ws[0]) < 3)


def norm(s: str) -> str:
    return " ".join(_word.findall(s.lower()))


class Engine:
    def __init__(self, db_path: str, termbase_ids: list[str], project_termbase: str | None = None):
        self.db_path = db_path
        self.con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, check_same_thread=False)
        self.lock = threading.Lock()
        self.termbases = self._load_termbases(termbase_ids)
        self.project_termbase = None
        if project_termbase:
            self.project_termbase = next((t for t in self.termbases
                                          if str(t["id"]) == project_termbase or t["name"] == project_termbase), None)
        self.terms: dict[str, dict] = {}          # norm(source) -> {src, tgts, kind, pair}
        self.max_words = 1
        for tb in self.termbases:
            self._load_terms(tb)
        self._load_pending()
        tables = {r[0] for r in self.con.execute("select name from sqlite_master where type='table'")}
        self.has_tm = {"translation_units", "translation_units_fts"} <= tables
        self.tm_units = self.con.execute("select count(*) from translation_units").fetchone()[0] if self.has_tm else 0
        self.tms = self._load_tms() if self.has_tm else []

    # ---- loading
    def _load_termbases(self, ids):
        if ids == ["all"]:
            rows = self.con.execute("select id, name, source_lang, target_lang from termbases order by id").fetchall()
        else:
            rows = self.con.execute(
                "select id, name, source_lang, target_lang from termbases where id in (%s)"
                % ",".join("?" * len(ids)), ids).fetchall()
        return [{"id": r[0], "name": r[1], "src": (r[2] or "").lower(), "tgt": (r[3] or "").lower(), "count": 0} for r in rows]

    def _add_term(self, src, tgt, kind, pair, tb_id=None):
        key = norm(src)
        if not key or trivial(key):
            return
        e = self.terms.setdefault(key, {"src": src, "tgts": [], "kind": kind, "pair": pair})
        if tgt and not any(t["t"] == tgt and t["tb"] == tb_id for t in e["tgts"]):
            e["tgts"].append({"t": tgt, "tb": tb_id})
        if kind == "nt":
            e["kind"] = "nt"
        self.max_words = max(self.max_words, key.count(" ") + 1)

    def _add_pair(self, src, tgt, kind, pair, tb_id=None):
        """A termbase entry serves both directions: BEIJER is stored
        English-first and is used on Dutch-to-English jobs, as Studio does."""
        self._add_term(src, tgt, kind, pair, tb_id)
        if tgt:
            self._add_term(tgt, src, kind, (pair[1], pair[0]), tb_id)

    def _load_terms(self, tb):
        rows = self.con.execute(
            "select source_term, target_term, is_nontranslatable, synonyms, source_lang, target_lang "
            "from termbase_terms where termbase_id = ?", (str(tb["id"]),)).fetchall()
        for src, tgt, nt, syn, sl, tl in rows:
            if not src:
                continue
            # the term's own languages win; a termbase's header can be wrong or mixed
            pair = ((sl or tb["src"]).lower()[:2], (tl or tb["tgt"]).lower()[:2])
            self._add_pair(src, tgt or "", "nt" if nt else "", pair, str(tb["id"]))
            for s in (syn or "").split(";"):
                if s.strip():
                    self._add_term(src, s.strip(), "", pair)
        tb["count"] = len(rows)

    def _load_tms(self):
        counts = dict(self.con.execute("select tm_id, count(*) from translation_units group by tm_id").fetchall())
        rows = self.con.execute("select id, name, tm_id, source_lang, target_lang from translation_memories order by name").fetchall()
        return [{"id": r[0], "name": r[1], "tm_id": r[2], "src": (r[3] or "").lower()[:2], "tgt": (r[4] or "").lower()[:2],
                 "count": counts.get(r[2], 0)} for r in rows]

    def _load_pending(self):
        if os.path.exists(PENDING):
            for t in json.load(open(PENDING, encoding="utf-8")):
                self._add_pair(t["src"], t["tgt"], "", (t.get("src_lang", ""), t.get("tgt_lang", "")))

    # ---- queries
    def status(self):
        return {"ok": True, "db": os.path.basename(self.db_path), "termbases": self.termbases, "tms": self.tms,
                "project_termbase": self.project_termbase, "terms": len(self.terms), "tm_units": self.tm_units}

    def term_hits(self, text: str, src: str, tgt: str, tbs: set | None = None):
        """`tbs`: termbase ids switched on for this document; None means all."""
        words = _word.findall(text.lower())
        seen, out = set(), []
        for n in range(min(self.max_words, len(words)), 0, -1):
            for i in range(len(words) - n + 1):
                key = " ".join(words[i:i + n])
                e = self.terms.get(key)
                if not e or key in seen or (e["pair"][0] and e["pair"] != (src, tgt)):
                    continue
                tgts = [t["t"] for t in e["tgts"] if tbs is None or t["tb"] is None or t["tb"] in tbs]
                if not tgts:
                    continue
                seen.add(key)
                out.append({"src": e["src"], "tgts": tgts, "kind": e["kind"]})
        return out

    def tm_candidates(self, text: str, src: str, tgt: str, limit: int = 40, tms: list | None = None):
        words = sorted({w for w in _word.findall(text.lower()) if len(w) >= 4}, key=len, reverse=True)[:8]
        if not words or not self.has_tm:
            return []
        q = " OR ".join('"' + w.replace('"', "") + '"' for w in words)
        sql = ("select t.source_text, t.target_text, t.tm_id from translation_units_fts f "
               "join translation_units t on t.id = f.rowid "
               "where translation_units_fts match ? and lower(substr(t.source_lang,1,2)) = ? "
               "and lower(substr(t.target_lang,1,2)) = ?")
        args = [q, src, tgt]
        if tms is not None:
            sql += " and t.tm_id in (%s)" % ",".join("?" * len(tms))
            args += tms
        sql += " limit ?"
        args.append(limit)
        with self.lock:
            rows = self.con.execute(sql, args).fetchall()
        return [{"source": s, "target": t, "name": name} for s, t, name in rows]

    def add_term(self, src: str, tgt: str, src_lang: str, tgt_lang: str, termbase_id=None):
        """Insert one term the way the Trados plugin does (same columns), into
        the given termbase or the project termbase. Own short-lived read-write
        connection; WAL mode lets Studio keep the database open meanwhile."""
        tb = next((t for t in self.termbases if str(t["id"]) == str(termbase_id)), None) if termbase_id else self.project_termbase
        if tb is None:
            raise ValueError("no destination termbase")
        rw = sqlite3.connect(self.db_path, timeout=5)
        try:
            with rw:
                rw.execute(
                    "INSERT INTO termbase_terms (source_term, target_term, termbase_id, source_lang, target_lang, "
                    "definition, domain, notes, forbidden, case_sensitive, is_nontranslatable, term_uuid, "
                    "source_abbreviation, target_abbreviation, url, client, project, part_of_speech, context) "
                    "VALUES (?, ?, ?, ?, ?, '', '', '', 0, 0, 0, ?, '', '', '', '', '', '', '')",
                    (src.strip(), tgt.strip(), int(tb["id"]), src_lang, tgt_lang, str(uuid.uuid4())))
        finally:
            rw.close()
        tb["count"] = tb.get("count", 0) + 1
        self._add_pair(src, tgt, "", (src_lang[:2], tgt_lang[:2]), str(tb["id"]))
        return tb


class Handler(SimpleHTTPRequestHandler):
    engine: Engine = None

    def end_headers(self):                        # never let Word's browser cache the pane
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, fmt, *args):            # quiet, except errors
        if args and str(args[1]).startswith(("4", "5")):
            super().log_message(fmt, *args)

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/pane":
            # The manifest points here. Redirecting to a version-stamped address
            # defeats Office's add-in page cache, which ignores no-store.
            stamp = int(os.path.getmtime(os.path.join(STATIC, "index.html")))
            self.send_response(302)
            self.send_header("Location", "/index.html?v=%d" % stamp)
            self.end_headers()
            return
        if not u.path.startswith("/api/"):
            return super().do_GET()
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        e = self.engine
        try:
            if u.path == "/api/status":
                return self._json(e.status())
            tbs = set(x for x in q["tbs"].split(",") if x) if "tbs" in q else None
            tms = [x for x in q["tms"].split(",") if x] if "tms" in q else None
            if u.path == "/api/terms":
                return self._json(e.term_hits(q.get("text", ""), q.get("src", "").lower(), q.get("tgt", "").lower(), tbs))
            if u.path == "/api/tm":
                return self._json(e.tm_candidates(q.get("text", ""), q.get("src", "").lower(), q.get("tgt", "").lower(), tms=tms))
        except Exception as ex:                       # noqa: BLE001
            return self._json({"error": str(ex)}, 500)
        self._json({"error": "unknown endpoint"}, 404)

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        data = json.loads(self.rfile.read(n) or b"{}")
        if u.path == "/api/terms":
            src, tgt = (data.get("src") or "").strip(), (data.get("tgt") or "").strip()
            if not src or not tgt:
                return self._json({"error": "src and tgt required"}, 400)
            try:
                tb = self.engine.add_term(src, tgt, (data.get("src_lang") or "").lower(), (data.get("tgt_lang") or "").lower(),
                                          data.get("termbase_id"))
            except Exception as ex:                   # noqa: BLE001
                return self._json({"error": str(ex)}, 500)
            return self._json({"ok": True, "termbase": tb["name"]})
        self._json({"error": "unknown endpoint"}, 404)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", required=True)
    ap.add_argument("--termbases", default="all", help="comma-separated termbase ids, or all")
    ap.add_argument("--project-termbase", default=None, help="id or name of the termbase new terms go to")
    ap.add_argument("--port", type=int, default=3000)
    a = ap.parse_args()
    Handler.engine = Engine(a.db, [x.strip() for x in a.termbases.split(",") if x.strip()], a.project_termbase)
    if a.project_termbase and Handler.engine.project_termbase is None:
        raise SystemExit("project termbase not found: " + a.project_termbase)
    st = Handler.engine.status()
    print("engine: %s | %d termbases | %d terms | %d TM units | new terms -> %s"
          % (st["db"], len(st["termbases"]), st["terms"], st["tm_units"],
             st["project_termbase"]["name"] if st["project_termbase"] else "(none)"))
    srv = ThreadingHTTPServer(("127.0.0.1", a.port), partial(Handler, directory=STATIC))
    print("serving the pane on http://localhost:%d/index.html" % a.port)
    srv.serve_forever()


if __name__ == "__main__":
    main()
