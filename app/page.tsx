import type { Metadata } from "next";
import { MeetingNotesApp } from "./MeetingNotesApp";

export const metadata: Metadata = {
  title: "회의록 요약 | Meeting Notes",
  description:
    "Your first version will appear here automatically when it’s ready.",
};

export default function Home() {
  return <MeetingNotesApp />;
}
