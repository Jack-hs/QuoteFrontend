import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// export default defineConfig({
//   plugins: [react()],
//   server: {
//     proxy: {
//       '/api': {
//         target: 'http://localhost:5129',
//         changeOrigin: true,
//         secure: false  // HTTP！
//       }
//     }
//   }
// })
export default defineConfig({
  plugins: [react()],
  base: '/crispy-train/',  // ← 你的 repo 名稱 + /
})