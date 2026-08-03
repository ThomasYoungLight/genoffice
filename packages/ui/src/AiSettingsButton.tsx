import type { AiSettings } from '@genoffice/ai-provider'
import type { Lang } from '@genoffice/i18n'
import { tAiSettings } from './ai-settings-strings'
import { IconSettings } from './icons'
import { useAiSettingsDialog, type AiSettingsHost } from './use-ai-settings-dialog'

export interface AiSettingsButtonProps {
  host: AiSettingsHost
  lang: Lang
  /** the freshly saved settings, so the host can keep its own copy in sync */
  onSaved?: ((settings: AiSettings) => void) | undefined
  /** the host app's own header-button class, so the gear matches its neighbours */
  className?: string | undefined
  iconSize?: number | undefined
}

/** Gear button that opens the shared provider dialog. */
export function AiSettingsButton({
  host,
  lang,
  onSaved,
  className,
  iconSize = 16,
}: AiSettingsButtonProps) {
  const { open, dialog } = useAiSettingsDialog(host, lang, onSaved)

  return (
    <>
      <button
        className={className}
        onClick={open}
        title={tAiSettings(lang, 'title')}
        aria-label={tAiSettings(lang, 'title')}
      >
        <IconSettings size={iconSize} />
      </button>
      {dialog}
    </>
  )
}

export type { AiSettingsHost }
