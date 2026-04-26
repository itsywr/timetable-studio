import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Static-only build for GitHub Pages.
// Repo: https://github.com/itsywr/timetable-studio  →  base /timetable-studio/
export default defineConfig({
  plugins: [react()],
  base: '/timetable-studio/',
  server: {
    host: true,
    port: 5173,
  },
})
