import { useState, type ReactNode } from 'react'
import type {
  AiProviderProbeRequest,
  AiSettings,
  CliStatus,
  ModelListResult,
  ProviderProbeResult,
} from '@genoffice/ai-provider'
import type { Lang } from '@genoffice/i18n'
import { AiSettingsDialog } from './AiSettingsDialog'

/**
 * The app's IPC surface, adapted. Each app names these differently
 * (`window.desktop.getAiSettings`, `window.slidesApi.getAiSettings`, …), so
 * the caller passes the two or four functions rather than this package
 * guessing at a global.
 */
export interface GensparkStatus {
  loggedIn: boolean
  email?: string | undefined
}

export interface AiSettingsHost {
  getAiSettings(): Promise<AiSettings>
  setAiSettings(settings: AiSettings): Promise<void>
  /** genspark login state, when the app exposes it */
  gskStatus?: (() => Promise<GensparkStatus>) | undefined
  /** start the genspark browser login, when the app exposes it */
  gskLogin?: (() => void) | undefined
  /** verify a key/model/endpoint before the user relies on it */
  aiTestProvider?: ((request: AiProviderProbeRequest) => Promise<ProviderProbeResult>) | undefined
  /** re-read the provider's model catalogue */
  aiListModels?: ((request: AiProviderProbeRequest) => Promise<ModelListResult>) | undefined
  /** whether a local agent CLI backend is installed */
  aiCliStatus?: ((provider: string) => Promise<CliStatus>) | undefined
}

/**
 * Mounts the provider dialog on demand. Shared by the header gear and the
 * composer's model picker so both open it the same way.
 *
 * Settings are re-read on open rather than taken from the caller's copy:
 * another window may have changed the provider since, and all windows share
 * one settings file.
 */
export function useAiSettingsDialog(
  host: AiSettingsHost,
  lang: Lang,
  onSaved?: ((settings: AiSettings) => void) | undefined,
): { open: () => void; dialog: ReactNode } {
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [genspark, setGenspark] = useState<GensparkStatus | null>(null)

  const open = () => {
    void host.getAiSettings().then(setSettings)
    if (host.gskStatus) void host.gskStatus().then(setGenspark, () => setGenspark(null))
  }

  const save = (next: AiSettings) => {
    setSettings(null)
    void host.setAiSettings(next).then(() => onSaved?.(next))
  }

  return {
    open,
    dialog: settings ? (
      <AiSettingsDialog
        settings={settings}
        lang={lang}
        onSave={save}
        onClose={() => setSettings(null)}
        onTest={host.aiTestProvider}
        onListModels={host.aiListModels}
        onCliStatus={host.aiCliStatus}
        genspark={genspark && host.gskLogin ? { ...genspark, onSignIn: host.gskLogin } : undefined}
      />
    ) : null,
  }
}
