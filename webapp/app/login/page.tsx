"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";

export default function Login() {
  const [passcode, setPasscode] = useState("");
  const [err, setErr] = useState(false);
  const router = useRouter();

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    const res = await fetch("/api/auth", { method: "POST", body: JSON.stringify({ passcode }) });
    if (res.ok) router.push("/");
    else setErr(true);
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-950 text-slate-100">
      <form onSubmit={submit} className="w-80 rounded-xl border border-slate-800 bg-slate-900 p-6 shadow-xl">
        <h1 className="mb-1 text-lg font-semibold">SecretaryBench</h1>
        <p className="mb-4 text-sm text-slate-400">Corpus authoring — enter the club passcode.</p>
        <input type="password" value={passcode} onChange={(e) => setPasscode(e.target.value)} autoFocus
          placeholder="passcode" className="mb-3 w-full rounded-md border border-slate-700 bg-slate-800 px-3 py-2 text-sm outline-none focus:border-sky-500" />
        {err && <p className="mb-3 text-sm text-rose-400">Wrong passcode.</p>}
        <button type="submit" className="w-full rounded-md bg-sky-600 px-3 py-2 text-sm font-medium hover:bg-sky-500">Enter</button>
      </form>
    </div>
  );
}
