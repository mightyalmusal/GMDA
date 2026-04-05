import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'
import { createRequire } from 'node:module'

function localMetaInsightsApi() {
  const fnPath = resolve(process.cwd(), 'server/meta-insights.js')
  const cjsRequire = createRequire(import.meta.url)

  return {
    name: 'local-meta-insights-api',
    configureServer(server) {
      server.middlewares.use('/api/meta-insights', async (req, res) => {
        try {
          let rawBody = ''
          await new Promise((resolveBody, rejectBody) => {
            req.on('data', chunk => {
              rawBody += chunk
            })
            req.on('end', resolveBody)
            req.on('error', rejectBody)
          })

          delete cjsRequire.cache[fnPath]
          const mod = cjsRequire(fnPath)
          const handler = mod.handler || mod.default?.handler || mod.default

          if (typeof handler !== 'function') {
            throw new Error('meta-insights handler not found')
          }

          const response = await handler({
            httpMethod: req.method || 'GET',
            headers: req.headers || {},
            body: rawBody || '{}',
          })

          res.statusCode = response?.statusCode || 200
          const headers = response?.headers || { 'Content-Type': 'application/json' }
          Object.entries(headers).forEach(([k, v]) => res.setHeader(k, v))
          res.end(response?.body || '')
        } catch (err) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: err?.message || 'Internal Server Error' }))
        }
      })
    },
  }
}

export default defineConfig({
  plugins: [react(), localMetaInsightsApi()],
})
