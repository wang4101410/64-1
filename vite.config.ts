import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // IMPORTANT: Replace '/iso-14064-generator/' with your actual GitHub repository name if deploying to GitHub Pages.
  // If deploying to the root of a domain (e.g., netlify/vercel), you can remove this line or set it to '/'.
  base: './', 
})