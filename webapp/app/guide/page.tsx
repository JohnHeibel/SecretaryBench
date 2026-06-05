import Link from "next/link";

// A plain, friendly walkthrough. Past the four-step loop, everything lives in collapsed
// <details> cards so a newcomer sees the whole map and opens only what they need. This is
// the page the editor's "How this works" primer links to.

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

      <p className="mb-3 text-sm leading-relaxed text-slate-300">
        Here is the whole idea. You write a pretend email to a busy executive&apos;s <strong>AI assistant</strong>. Then you
        tell us the one right thing the assistant should do with it. We use your answer to grade real AI assistants.
      </p>
      <p className="mb-5 text-sm leading-relaxed text-slate-300">
        You can&apos;t break anything. If a test is not valid, the app blocks the export until it is fixed, so feel free to
        poke around.
      </p>

      {/* The core loop, always visible: the four things you actually do. */}
      <ol className="mb-6 grid gap-2 sm:grid-cols-2">
        {[
          ["1. Write the email", "Who it is from, a subject, and a few lines of body."],
          ["2. Add any dates", "Click + insert date and pick the pieces. The real day shows up as you build it."],
          ["3. Fill in the answer key", "Say what the assistant should do, or tick “this email needs no action.”"],
          ["4. Get the green bar", "Green means a perfect assistant could solve it. That is done."],
        ].map(([h, d]) => (
          <li key={h} className="rounded-lg border border-slate-800 bg-slate-900/40 p-3">
            <div className="text-sm font-medium text-sky-300">{h}</div>
            <div className="mt-0.5 text-xs text-slate-400">{d}</div>
          </li>
        ))}
      </ol>

      <p className="mb-2 text-xs uppercase tracking-wide text-slate-500">Open anything you want to understand better</p>
      <div className="space-y-2.5">
        <Card title="What makes an email easy or hard" teaser="the three shapes">
          <p>Every email you write is one of these three shapes. <strong>All three are graded, and all three matter.</strong></p>
          <ul className="space-y-1.5">
            <li><span className="font-medium text-emerald-300">Everything is in this email.</span> The assistant has all it needs right here. This is the easy kind.</li>
            <li><span className="font-medium text-rose-300">It needs an earlier email.</span> To answer this one, the assistant has to go find a fact from an email that came before. This is the interesting kind, and we call it a <strong>needle</strong>.</li>
            <li><span className="font-medium text-slate-300">Nothing to do.</span> The email looks like it might need action, but it does not. The assistant passes by doing nothing. Acting on these is a really common mistake, so write plenty of them.</li>
          </ul>
        </Card>

        <Card title="Writing a date" teaser="build it, see the real day">
          <p>Our dates are <em>relative</em>, like &ldquo;5 days after this email arrives&rdquo; or &ldquo;two weeks after the signing.&rdquo; So you don&apos;t click a day on a calendar. You <strong>build</strong> the date from a few choices: pick a starting day, then optionally <strong>shift it</strong> forward or back, like &ldquo;+2 weeks.&rdquo; A shift is just plain calendar math on the starting day, so &ldquo;the kickoff, +2 weeks&rdquo; lands two weeks after the kickoff.</p>
          <p>In the answer key the builder is right there. In the email body, click <span className={chip}>+ insert date</span>. Either way the real day shows up as you build it, like <span className="text-emerald-400">→ Monday, Aug 17, 2026</span>.</p>
          <p>The same date you build fills the body <em>and</em> the answer key, so they can never disagree. If you would rather type the date out yourself, you can, and it is checked live.</p>
        </Card>

        <Card title="The answer key" teaser="the one right thing to do">
          <p>The answer key sits under the body. It is your private answer, and the assistant never sees it.</p>
          <p>If the email needs no action, tick <span className={chip}>this email needs no action</span>. Now it is graded on the assistant doing nothing.</p>
          <p>Otherwise pick an action: <strong>create</strong> an event or to-do, <strong>move</strong> one, or <strong>cancel</strong> one. Give it a short name like <span className={chip}>kickoff</span>. That name is how a later email can refer back to it, and it is the word we look for in the assistant&apos;s calendar. A to-do gets a due date, where <strong>on or before</strong> means any day up to the deadline counts.</p>
        </Card>

        <Card title="Build a needle (the fun part)" teaser="two linked emails">
          <p>A needle is two emails that point at each other:</p>
          <ol className="ml-4 list-decimal space-y-1">
            <li>In <strong>email A&apos;s body</strong>, insert a date and tick <span className={chip}>other emails can refer to this date as</span>, then name it (like <span className={chip}>signing</span>).</li>
            <li>In <strong>email B&apos;s answer key</strong>, build the date from <span className={chip}>@signing</span> and add an offset like &ldquo;+ 1 week&rdquo; if you want.</li>
            <li>The link from A to B is added for you.</li>
          </ol>
          <p>Now the assistant has to find email A to answer email B. The more filler you bury between them, the harder it gets. That gap is the whole point.</p>
        </Card>

        <Card title="Three examples, easy to hard" teaser="see them">
          <div className="space-y-1">
            <div className="text-xs font-semibold text-emerald-300">Easy: everything is in the email</div>
            <Mail from="Sam (Office Manager)" subject="Quick intro sync">{"Can you put a 30-minute intro sync on the calendar for next Thursday?"}</Mail>
            <p className="text-xs text-slate-400">→ create an event named <span className={chip}>intro</span> on next Thursday. The date and the action are both right there.</p>
          </div>
          <div className="space-y-1">
            <div className="text-xs font-semibold text-amber-300">Harder: it needs the email before it</div>
            <Mail from="Lee (Partnerships)" subject="Partner visit">{"The partner visit is confirmed for one week from today."}</Mail>
            <Mail from="Lee (Partnerships)" subject="Welcome dinner">{"Let's host a welcome dinner the evening before the partner visit."}</Mail>
            <p className="text-xs text-slate-400">→ create the dinner the day <em>before</em> the visit. &ldquo;Before the visit&rdquo; only means something if you remember the first email.</p>
          </div>
          <div className="space-y-1">
            <div className="text-xs font-semibold text-rose-300">Hardest: the task nobody spelled out</div>
            <Mail from="Priya (Counsel)" subject="Apex regulatory note">{"Legal flagged that the HSR filing window closes thirty days after close, and we cannot miss it."}</Mail>
            <p className="text-xs text-slate-400">→ add a to-do <span className={chip}>HSR filing</span> due 30 days after close. Nobody said &ldquo;add a task.&rdquo; A weak assistant reads this as an FYI and does nothing. This is where strong and weak assistants split, so put your best effort here.</p>
          </div>
        </Card>

        <Card title="The green bar" teaser="your one signal">
          <p>The bar at the bottom runs the real benchmark checks as you type:</p>
          <ul className="space-y-1">
            <li><strong>Lint</strong> asks: would the benchmark even load this? (structure, names, links)</li>
            <li><strong>Oracle</strong> asks: could a <em>perfect</em> assistant actually solve every answer key?</li>
          </ul>
          <p>When both pass, you see <span className="text-emerald-400">ready to export</span>. You cannot export a broken one. That is the app protecting you.</p>
        </Card>

        <Card title="Words we use" teaser="storyline, cast, anchor…">
          <ul className="space-y-1">
            <li><span className={chip}>storyline</span> a group of related emails, one scenario. Make one per thread.</li>
            <li><span className={chip}>cast</span> the people in that storyline. Every storyline starts with <span className={chip}>CEO</span>, which is you, the person the assistant works for.</li>
            <li><span className={chip}>anchor</span> a date set in one email and reused in a later one, like &ldquo;the close date.&rdquo;</li>
            <li><span className={chip}>depends on</span> means &ldquo;this email comes after that one.&rdquo; The version with a deadline also passes a due date along.</li>
            <li><span className={chip}>needle</span> an email whose answer needs a fact from an earlier email.</li>
          </ul>
        </Card>

        <Card title="Things that trip people up" teaser="quick gotchas">
          <ul className="space-y-1">
            <li>Build dates with the date builder, or type the expression (it is checked live). The same date fills the body and the answer key, so they can&apos;t drift apart.</li>
            <li>You don&apos;t need to touch the ids. They name themselves from the subject.</li>
            <li>The <strong>To</strong> and <strong>Cc</strong> fields can hold several people. Tap a name to add or remove it. Cc is optional, and the assistant can see who is copied.</li>
            <li>Two different things in one storyline need different names. Don&apos;t call both &ldquo;review.&rdquo; Call them &ldquo;board review&rdquo; and &ldquo;client review.&rdquo;</li>
            <li>A date built from an anchor shows up amber (&ldquo;resolves when this email is sent&rdquo;). That is expected, not an error.</li>
          </ul>
        </Card>
      </div>

      <div className="mt-6 rounded-lg border border-slate-800 bg-slate-900/40 px-4 py-3 text-sm text-slate-400">
        That&apos;s it. Make a storyline, warm up with a couple of easy emails, add some tempting filler, then try a real needle.
        <Link href="/" className="ml-1 text-sky-400 hover:text-sky-300">Start writing →</Link>
      </div>
    </div>
  );
}
