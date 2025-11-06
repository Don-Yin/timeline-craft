import Link from "next/link";

export default function Landing() {
  return (
    <div className="flex min-h-screen w-full bg-zinc-50 font-sans dark:bg-black">
      <main className="flex min-h-screen w-full flex-col items-center gap-10 bg-white py-16 px-6 dark:bg-black sm:px-12">
        <div className="w-full max-w-3xl">
          <div className="mb-8 flex w-full items-center justify-between">
            <span className="text-xl font-semibold text-black dark:text-zinc-50">
              TimelineCraft
            </span>
          </div>

          <section
            className="group relative overflow-hidden rounded-2xl border border-zinc-200/60 bg-white/70 p-5 shadow-sm transition
                       hover:border-zinc-300/80 hover:shadow-md dark:border-white/10 dark:bg-black/30 dark:hover:border-white/20"
          >
            <div className="pointer-events-none absolute inset-0 -z-10 bg-gradient-to-b from-transparent to-black/[.02] dark:to-white/[.03]" />
            <div className="overflow-hidden rounded-xl border border-zinc-200/70 dark:border-white/10">
              <img
                src="https://github.com/don-yin/powerpoint-timeline/raw/cf5610f7db48a2f3e2fb747e2f197c5dbedd45e8/public/demo.gif"
                alt="TimelineCraft demo"
                className="h-auto w-full"
              />
            </div>
            <div className="mt-6 flex justify-center">
              <Link
                href="/operate"
                className="rounded-full bg-foreground px-6 py-3 text-background shadow-sm transition hover:opacity-95"
              >
                get started →
              </Link>
            </div>
          </section>
        </div>
      </main>
    </div>
  );
}


