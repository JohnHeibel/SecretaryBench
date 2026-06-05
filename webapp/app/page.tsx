import { Suspense } from "react";
import Workspace from "@/components/Workspace";

// Workspace reads ?node=<id> via useSearchParams (deep-link focus mode), which Next requires to sit
// inside a Suspense boundary so the rest of the route can still render statically.
export default function Home() {
  return (
    <Suspense fallback={null}>
      <Workspace />
    </Suspense>
  );
}
