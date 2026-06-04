import Link from "next/link";

// A friendly, scannable guide. Everything lives in collapsed <details> cards so a newcomer
// sees the whole map at a glance and opens only the one thing they need — never a wall of text.
// This is the single source of truth the editor's "How this works" primer links to.

function Card({ title, teaser, children }: { title: string; teaser: string; children: React.ReactNode }) {
  return (
    <details className="group rounded-xl border border-slate-800 bg-slate-900/40 transition-colors open:border-sky-900/70 open:bg-sky-950/20">
      <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3 [&::-webkit-details-marker]:hidden">
        <span className="text-sky-500 transition-transform group-open:rotate-90">▸</span>
        <span className="font-semibold text-slate-100">{title}</span>
        <span className="ml-auto truncate text-xs text-slate-500 group-open:hidden">{teaser}</span>
      </summary>
      <div className="space-y-3 border-t border-slate-800/80 px-4 py-3 text-sm leading-relaxed text-slate-300">{children}</div>
    </details>
  );
}

const chip = "rounded bg-slate-800 px-1.5 py-0.5 font-mono text-[12px] text-sky-300";

function Mail({ from, subject, children }: { from: string; subject: string; children: React.ReactNode }) {
  return (
    <div className="rounded-lg border border-slate-800 bg-slate-950/60 p-3 text-[13px]">
      <div className="text-slate-500"><span className="text-slate-400">From:</span> {from}</div>
      <div className="mb-1.5 text-slate-500"><span className="text-slate-400">Subject:</span> {subject}</div>
      <p className="text-slate-300">{children}</p>
    </div>
  );
}

export default function GuidePage() {
  return (
    <div className="mx-auto max-w-2xl px-5 py-8">
      <div className="mb-5 flex items-center justify-between">
        <h1 className="text-lg font-semibold text-slate-100">How to write a test</h1>
        <Link href="/" className="rounded-md border border-slate-700 px-3 py-1 text-xs text-slate-300 hover:bg-slate-800">← back to editor</Link>
      </div>

      <p className="mb-5 text-sm leading-relaxed text-slate-300">
        You write a fake email to a busy exec&apos;s <strong>AI assistant</strong>, then say what a <em>perfect</em> assistant
        should do with it. That&apos;s the whole job. You can&apos;t break anything — the app won&apos;t let a broken test out the door.
      </p>

      {/* The core loop, always visible — this is the "what you actually do". */}
      <ol className="mb-6 grid gap-2 sm:grid-cols-2">
        {[
          ["1 · Write the email", "Who it's from, a subject, a short body."],
          ["2 · Add any dates with the date builder", "Click + insert date and pick the pieces — the resolved day shows live."],
          ["3 · Fill the answer key", "Say what to do, or tick “needs no action.”"],
          ["4 · Watch the bar go green", "Green = a perfect assistant could solve it. Done."],
        ].map(([h, d]) => (
          <li key={h} className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
            <div className="text-sm font-medium text-sky-300">{h}</div>
            <div className="mt-0.5 text-xs text-slate-400">{d}</div>
          </li>
        ))}
      </ol>

      <p className="mb-2 text-xs uppercase tracking-wide text-slate-500">Open anything you want to understand better</p>
      <div className="space-y-2.5">
        <Card title="The 3 kinds of email" teaser="needle · test · filler">
          <p>Every email is one of these. <strong>All three are graded and all three matter.</strong></p>
          <ul className="space-y-1.5">
            <li><span className="font-medium text-rose-300">Needle</span> — the answer needs a fact from an <em>earlier</em> email. The AI has to remember/search back. The hard, juicy kind.</li>
            <li><span className="font-medium text-emerald-300">Test</span> — everything you need is in <em>this</em> email. Still fully graded, just no searching.</li>
            <li><span className="font-medium text-slate-300">Filler</span> — no action wanted. It&apos;s bait; the AI passes by <strong>doing nothing</strong>. (Over-acting on noise is a super common failure — write lots of these.)</li>
          </ul>
        </Card>

        <Card title="Writing a date" teaser="build it from choices — shown resolved live">
          <p>Our dates are <em>relative</em> (&ldquo;5 days after this email arrives,&rdquo; &ldquo;two weeks after the signing&rdquo;), so you don&apos;t pick a day on a calendar — you <strong>build</strong> it from a short sentence of choices: a starting point, then any offsets.</p>
          <p>In the answer key the builder is right there inline; in the email body, click <span className={chip}>+ insert date</span>. Either way the resolved day shows live (e.g. <span className="text-emerald-400">→ Monday, Aug 17, 2026</span>). The <em>same</em> built date drives the body <em>and</em> the answer key, so they can never disagree. You can also type the expression directly — it&apos;s checked against the real grader as you type.</p>
        </Card>

        <Card title="The answer key" teaser="what a perfect assistant should do">
          <p>Below the body. For a filler email, just tick <span className={chip}>this email needs no action</span> — it&apos;s now graded on the AI staying still.</p>
          <p>Otherwise add an action: <strong>create</strong> an event/to-do, <strong>move</strong>, or <strong>cancel</strong> one. Give it a short name like <span className={chip}>kickoff</span> — that name is how a later email refers back to it, and it&apos;s the word we look for in the AI&apos;s calendar title.</p>
        </Card>

        <Card title="Make a needle (the fun part)" teaser="2 emails + a link between them">
          <p>A needle is two emails that point at each other:</p>
          <ol className="ml-4 list-decimal space-y-1">
            <li>In <strong>Email A&apos;s body</strong>, insert a date and tick <span className={chip}>publish as anchor</span> → name it (e.g. <span className={chip}>signing</span>).</li>
            <li>In <strong>Email B&apos;s answer key</strong>, build the date using the <span className={chip}>@signing</span> block (plus an offset like &ldquo;+ 1 week&rdquo;).</li>
            <li>Add a <strong>date</strong> dependency edge from A → B.</li>
          </ol>
          <p>Now the AI must find A to answer B. The more filler we bury between them, the harder it gets — that gap is the whole point.</p>
        </Card>

        <Card title="Worked examples: easy → hard" teaser="see a real T1, T2, T3">
          <p className="text-xs text-slate-400">You tag each email T1 (easy) / T2 (medium) / T3 (hard) so we can score models by tier.</p>
          <div className="space-y-1">
            <div className="text-xs font-semibold text-emerald-300">T1 — easy (one email, nothing to remember)</div>
            <Mail from="Sam (Office Manager)" subject="Quick intro sync">Can you put a 30-min intro sync on the calendar for next Thursday?</Mail>
            <p className="text-xs text-slate-400">→ create event <span className={chip}>intro</span> on next Thursday. The date and the action are both right there.</p>
          </div>
          <div className="space-y-1">
            <div className="text-xs font-semibold text-amber-300">T2 — medium (one hop back)</div>
            <Mail from="Lee (Partnerships)" subject="Partner visit">The partner visit is confirmed for one week from today.</Mail>
            <Mail from="Lee (Partnerships)" subject="Welcome dinner">Let&apos;s host a welcome dinner the evening before the partner visit.</Mail>
            <p className="text-xs text-slate-400">→ create the dinner the day <em>before</em> the visit. &ldquo;Before the visit&rdquo; only makes sense if you remember the first email.</p>
          </div>
          <div className="space-y-1">
            <div className="text-xs font-semibold text-rose-300">T3 — hard (the implied task — the big one)</div>
            <Mail from="Priya (Counsel)" subject="Apex regulatory note">Legal flagged that the HSR filing window closes thirty days after close, and we cannot miss it.</Mail>
            <p className="text-xs text-slate-400">→ add a to-do <span className={chip}>HSR filing</span> due 30 days after close. Nobody said &ldquo;add a task&rdquo; — a weak model reads this as FYI and does nothing. That&apos;s where strong and weak models separate, so spend your best effort here.</p>
          </div>
        </Card>

        <Card title="The green bar" teaser="your single source of truth">
          <p>The bar at the bottom runs the <em>real</em> benchmark checks as you work:</p>
          <ul className="space-y-1">
            <li><strong>Lint</strong> — would the benchmark load this? (structure, ids, anchors)</li>
            <li><strong>Oracle</strong> — could a <em>perfect</em> assistant actually solve every answer key?</li>
          </ul>
          <p>Both green = <span className="text-emerald-400">ready to export</span>. You literally cannot export a broken one — that&apos;s the app protecting you.</p>
        </Card>

        <Card title="Words we use" teaser="node · cast · anchor · …">
          <ul className="space-y-1">
            <li><span className={chip}>storyline</span> a group of related emails — one scenario. Make one per thread. (Saved as a <em>node</em> file under the hood.)</li>
            <li><span className={chip}>cast</span> the people in that storyline. Every storyline starts with <span className={chip}>CEO → you</span> (the AI works for you).</li>
            <li><span className={chip}>anchor</span> a date set in one email and reused in a later one (e.g. &ldquo;the close date&rdquo;).</li>
            <li><span className={chip}>edge</span> &ldquo;this email comes after that one.&rdquo; A <em>date</em> edge also carries a deadline.</li>
            <li><span className={chip}>span</span> how far back the needed fact is — the main thing that makes an email hard.</li>
          </ul>
        </Card>

        <Card title="Quick gotchas" teaser="things that trip people up">
          <ul className="space-y-1">
            <li>Build dates with the date builder (or type the expression — it&apos;s checked live). The same date drives the body and the answer key, so they can&apos;t drift.</li>
            <li>Don&apos;t worry about ids — they auto-name from the subject. Leave them alone.</li>
            <li>The <strong>To</strong> field is one person.</li>
            <li>Two things in one storyline need different names (not two &ldquo;reviews&rdquo; — call them &ldquo;board&rdquo; and &ldquo;client&rdquo;).</li>
            <li>An anchor date previews as amber (&ldquo;resolves at serve time&rdquo;). That&apos;s expected, not an error.</li>
          </ul>
        </Card>
      </div>

      <div className="mt-6 rounded-lg border border-slate-800 bg-slate-900/40 px-4 py-3 text-sm text-slate-400">
        That&apos;s it. Make a storyline, warm up with a couple of easy tests, sprinkle in tempting filler, then try a real needle.
        <Link href="/" className="ml-1 text-sky-400 hover:text-sky-300">Start authoring →</Link>
      </div>
    </div>
  );
}
