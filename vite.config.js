import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// Wird die App über GitHub Actions gebaut, steht GITHUB_REPOSITORY automatisch
// zur Verfügung (z.B. "meinname/methodencurriculum-app"). Daraus wird der Pfad
// abgeleitet, unter dem GitHub Pages die Seite ausliefert
// (https://<user>.github.io/<repo>/) - beim lokalen "npm run dev" ist die
// Variable nicht gesetzt, dann bleibt die Basis "/".
const repoName = process.env.GITHUB_REPOSITORY?.split("/")[1];

export default defineConfig({
  plugins: [react(), tailwindcss()],
  base: repoName ? `/${repoName}/` : "/",
});
