'use client';

import Link from "next/link";
import { useState } from "react";
import { Play, Sparkles, Zap, Layers, ArrowRight, Download } from "lucide-react";

export default function Landing() {
  const [isVideoPlaying, setIsVideoPlaying] = useState(false);

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white overflow-x-hidden">
      {/* Navigation */}
      <nav className="fixed top-0 left-0 right-0 z-50 backdrop-blur-xl bg-black/60 border-b border-white/5">
        <div className="max-w-6xl mx-auto px-6 h-14 flex items-center justify-between">
          <span className="text-lg font-semibold tracking-tight">TimelineCraft</span>
          <Link
            href="/manage"
            className="text-sm text-zinc-400 hover:text-white transition-colors"
          >
            Open App →
          </Link>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="pt-32 pb-20 px-6">
        <div className="max-w-4xl mx-auto text-center">
          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-white/5 border border-white/10 text-xs text-zinc-400 mb-8">
            <Sparkles className="w-3.5 h-3.5 text-emerald-400" />
            automate your presentation workflow
          </div>

          <h1 className="text-5xl sm:text-7xl font-bold tracking-tight leading-[1.1] mb-6">
            sidebars with timeline
            <br />
            <span className="bg-gradient-to-r from-emerald-400 via-cyan-400 to-blue-400 bg-clip-text text-transparent">
              zero effort.
            </span>
          </h1>

          <p className="text-lg sm:text-xl text-zinc-400 max-w-2xl mx-auto mb-10 leading-relaxed">
            Transform your PowerPoint presentations with elegant timeline sidebars.
            Upload, configure, and download
          </p>

          <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
            <Link
              href="/manage"
              className="group flex items-center gap-2 px-8 py-4 rounded-full bg-white text-black font-medium text-sm hover:bg-zinc-200 transition-all"
            >
              Get Started Free
              <ArrowRight className="w-4 h-4 group-hover:translate-x-0.5 transition-transform" />
            </Link>
            <a
              href="/demo/after.pptx"
              download
              className="flex items-center gap-2 px-8 py-4 rounded-full border border-white/20 text-sm hover:bg-white/5 transition-all"
            >
              <Download className="w-4 h-4" />
              Download Example
            </a>
          </div>
        </div>
      </section>

      {/* Video Demo Section */}
      <section className="pb-20 px-6">
        <div className="max-w-5xl mx-auto">
          <div
            className="relative aspect-video rounded-2xl overflow-hidden bg-zinc-900 border border-white/10 shadow-2xl shadow-emerald-500/10 cursor-pointer group"
            onClick={() => setIsVideoPlaying(true)}
          >
            {!isVideoPlaying ? (
              <>
                <div className="absolute inset-0 bg-gradient-to-br from-emerald-500/10 via-transparent to-blue-500/10" />
                <div className="absolute inset-0 flex items-center justify-center">
                  <div className="w-20 h-20 rounded-full bg-white/10 backdrop-blur-sm flex items-center justify-center group-hover:bg-white/20 transition-all group-hover:scale-110">
                    <Play className="w-8 h-8 text-white ml-1" fill="white" />
                  </div>
                </div>
                <div className="absolute bottom-6 left-6 right-6 flex items-center justify-between">
                  <span className="text-sm text-zinc-400">Watch the demo</span>
                  <span className="text-xs text-zinc-500">0:30</span>
                </div>
              </>
            ) : (
              <video
                autoPlay
                controls
                playsInline
                className="w-full h-full object-cover"
              >
                <source src="/demo/example.mp4" type="video/mp4" />
                <source src="/demo/example.mov" type="video/quicktime" />
              </video>
            )}
          </div>
        </div>
      </section>

      {/* Before/After Section */}
      <section className="py-20 px-6 bg-gradient-to-b from-transparent via-emerald-950/20 to-transparent">
        <div className="max-w-6xl mx-auto">
          <div className="text-center mb-16">
            <h2 className="text-3xl sm:text-4xl font-bold tracking-tight mb-4">
              see the transformation
            </h2>
            <p className="text-zinc-400 max-w-xl mx-auto">
              your plain presentations become professionally structured with navigable timeline sidebars.
            </p>
          </div>

          <div className="grid md:grid-cols-2 gap-8">
            {/* Before */}
            <div className="group">
              <div className="relative rounded-xl overflow-hidden bg-zinc-900 border border-white/10 p-1">
                <div className="absolute top-4 left-4 z-10 px-3 py-1 rounded-full bg-zinc-800/80 backdrop-blur-sm text-xs font-medium">
                  Before
                </div>
                <div className="aspect-[16/10] rounded-lg bg-zinc-800 flex items-center justify-center">
                  <div className="text-center p-8">
                    <Layers className="w-12 h-12 text-zinc-600 mx-auto mb-4" />
                    <p className="text-zinc-500 text-sm">Standard presentation</p>
                    <p className="text-zinc-600 text-xs mt-1">No navigation structure</p>
                  </div>
                </div>
              </div>
              <a
                href="/demo/before.pptx"
                download
                className="mt-3 inline-flex items-center gap-2 text-xs text-zinc-500 hover:text-zinc-300 transition-colors"
              >
                <Download className="w-3.5 h-3.5" />
                Download before.pptx
              </a>
            </div>

            {/* After */}
            <div className="group">
              <div className="relative rounded-xl overflow-hidden bg-zinc-900 border border-emerald-500/30 p-1 shadow-lg shadow-emerald-500/10">
                <div className="absolute top-4 left-4 z-10 px-3 py-1 rounded-full bg-emerald-500/20 backdrop-blur-sm text-xs font-medium text-emerald-400">
                  After
                </div>
                <div className="aspect-[16/10] rounded-lg bg-gradient-to-br from-zinc-800 to-zinc-900 flex items-center justify-center relative overflow-hidden">
                  <div className="absolute left-0 top-0 bottom-0 w-16 bg-gradient-to-r from-emerald-600/40 to-emerald-600/20 flex flex-col items-center justify-center gap-2 py-4">
                    {['Intro', 'Data', 'Results', 'End'].map((label, i) => (
                      <div key={label} className={`w-12 h-6 rounded text-[8px] flex items-center justify-center ${i === 1 ? 'bg-zinc-900/80 text-white' : 'text-white/60'}`}>
                        {label}
                      </div>
                    ))}
                  </div>
                  <div className="text-center p-8 ml-8">
                    <Sparkles className="w-12 h-12 text-emerald-400 mx-auto mb-4" />
                    <p className="text-zinc-300 text-sm">With timeline sidebar</p>
                    <p className="text-emerald-400/60 text-xs mt-1">Clear navigation & structure</p>
                  </div>
                </div>
              </div>
              <a
                href="/demo/after.pptx"
                download
                className="mt-3 inline-flex items-center gap-2 text-xs text-emerald-400/70 hover:text-emerald-400 transition-colors"
              >
                <Download className="w-3.5 h-3.5" />
                Download after.pptx
              </a>
            </div>
          </div>
        </div>
      </section>

      {/* Features Section */}
      <section className="py-20 px-6">
        <div className="max-w-5xl mx-auto">
          <div className="grid sm:grid-cols-3 gap-6">
            {[
              {
                icon: Zap,
                title: "Lightning Fast",
                description: "Process presentations in seconds, not hours. Our optimized pipeline handles even large decks effortlessly.",
                color: "text-yellow-400",
              },
              {
                icon: Layers,
                title: "Smart Sections",
                description: "Automatically organize slides into logical sections with drag-and-drop ease.",
                color: "text-blue-400",
              },
              {
                icon: Sparkles,
                title: "Smooth Transitions",
                description: "Built-in morph transitions create seamless, professional animations between slides.",
                color: "text-emerald-400",
              },
            ].map((feature) => (
              <div
                key={feature.title}
                className="p-6 rounded-2xl bg-white/[0.02] border border-white/5 hover:border-white/10 transition-all"
              >
                <feature.icon className={`w-8 h-8 ${feature.color} mb-4`} />
                <h3 className="font-semibold mb-2">{feature.title}</h3>
                <p className="text-sm text-zinc-500 leading-relaxed">{feature.description}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-20 px-6">
        <div className="max-w-3xl mx-auto text-center">
          <h2 className="text-3xl sm:text-4xl font-bold tracking-tight mb-4">
            Ready to upgrade your presentations?
          </h2>
          <p className="text-zinc-400 mb-8">
            Join thousands of professionals creating stunning presentations with TimelineCraft.
          </p>
          <Link
            href="/manage"
            className="inline-flex items-center gap-2 px-8 py-4 rounded-full bg-gradient-to-r from-emerald-500 to-cyan-500 text-black font-medium text-sm hover:opacity-90 transition-all"
          >
            Start Creating Now
            <ArrowRight className="w-4 h-4" />
          </Link>
        </div>
      </section>

      {/* Footer */}
      <footer className="py-8 px-6 border-t border-white/5">
        <div className="max-w-6xl mx-auto flex items-center justify-between text-xs text-zinc-600">
          <span>© 2026 TimelineCraft</span>
          <span>Built with ♥</span>
        </div>
      </footer>
    </div>
  );
}
