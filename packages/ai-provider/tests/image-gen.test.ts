import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  DEFAULT_IMAGE_MODELS,
  generateProviderImage,
  imageModelFor,
  providerGeneratesImages,
} from '../src/images'
import { errorResponse, jsonResponse } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const config = { apiKey: 'k', model: 'gpt-5.6-sol' }
const bodyOf = (call: unknown) => JSON.parse((call as { body: string }).body) as Record<string, any>

function stub(...responses: Response[]) {
  const fetchMock = vi.fn()
  for (const response of responses) fetchMock.mockResolvedValueOnce(response)
  vi.stubGlobal('fetch', fetchMock)
  return fetchMock
}

describe('imageModelFor', () => {
  it('prefers what the settings file names over the build default', () => {
    expect(imageModelFor('openai', 'gpt-image-3-preview')).toBe('gpt-image-3-preview')
    expect(imageModelFor('openai', '  ')).toBe(DEFAULT_IMAGE_MODELS.openai)
    expect(imageModelFor('openai', undefined)).toBe(DEFAULT_IMAGE_MODELS.openai)
  })

  it('has nothing to offer a provider with no image endpoint', () => {
    expect(providerGeneratesImages('openai')).toBe(true)
    expect(providerGeneratesImages('anthropic')).toBe(false)
    expect(providerGeneratesImages('claude-cli')).toBe(false)
    expect(imageModelFor('anthropic', undefined)).toBe('')
  })
})

describe('generateProviderImage', () => {
  it('posts the prompt and returns the bytes, never a hosted URL', async () => {
    const fetchMock = stub(jsonResponse({ data: [{ b64_json: 'iVBORw0KGgo=' }] }))
    const result = await generateProviderImage('openai', config, 'gpt-image-2', {
      prompt: 'a cutaway diagram of a jet engine',
      size: '1536x1024',
      quality: 'high',
    })

    expect(fetchMock.mock.calls[0]![0]).toBe('https://api.openai.com/v1/images/generations')
    expect(bodyOf(fetchMock.mock.calls[0]![1])).toEqual({
      model: 'gpt-image-2',
      prompt: 'a cutaway diagram of a jet engine',
      n: 1,
      size: '1536x1024',
      quality: 'high',
    })
    expect(result).toEqual({
      ok: true,
      base64: 'iVBORw0KGgo=',
      mime: 'image/png',
      model: 'gpt-image-2',
    })
  })

  it('asks for a transparent background as png', async () => {
    const fetchMock = stub(jsonResponse({ data: [{ b64_json: 'AAA' }] }))
    await generateProviderImage('openai', config, 'gpt-image-2', {
      prompt: 'a logo',
      transparent: true,
    })
    const body = bodyOf(fetchMock.mock.calls[0]![1])
    expect(body.background).toBe('transparent')
    expect(body.output_format).toBe('png')
  })

  it('retries with the minimum every image API accepts when an option is rejected', async () => {
    const fetchMock = stub(
      errorResponse(400, '{"error":{"message":"Unknown parameter: \'background\'."}}'),
      jsonResponse({ data: [{ b64_json: 'AAA' }] }),
    )
    const result = await generateProviderImage('openai', config, 'dall-e-3', {
      prompt: 'a logo',
      size: '1024x1024',
      quality: 'high',
      transparent: true,
    })
    expect(result.ok).toBe(true)
    expect(bodyOf(fetchMock.mock.calls[1]![1])).toEqual({
      model: 'dall-e-3',
      prompt: 'a logo',
      n: 1,
      size: '1024x1024',
    })
  })

  it('downloads the image when the endpoint hands back a URL instead', async () => {
    const png = new Response(new Uint8Array([1, 2, 3]), {
      status: 200,
      headers: { 'content-type': 'image/png' },
    })
    stub(jsonResponse({ data: [{ url: 'https://cdn.example/img.png' }] }), png)
    const result = await generateProviderImage('openai', config, 'dall-e-3', { prompt: 'x' })
    expect(result.ok).toBe(true)
    expect(result.base64).toBe(Buffer.from([1, 2, 3]).toString('base64'))
  })

  it('explains a rejected key and a missing model rather than echoing a status code', async () => {
    stub(errorResponse(401, 'bad key'))
    expect(
      (await generateProviderImage('openai', config, 'gpt-image-2', { prompt: 'x' })).error,
    ).toContain('the API key was rejected')

    stub(errorResponse(404, 'no such model'))
    expect(
      (await generateProviderImage('openai', config, 'gpt-image-9', { prompt: 'x' })).error,
    ).toContain('the image model was not found for this account')
  })

  it('refuses before spending a request when it has nothing to send', async () => {
    const fetchMock = stub()
    expect(
      (await generateProviderImage('anthropic', config, 'gpt-image-2', { prompt: 'x' })).error,
    ).toContain('no image generation endpoint')
    expect(
      (
        await generateProviderImage('openai', { apiKey: '', model: '' }, 'gpt-image-2', {
          prompt: 'x',
        })
      ).error,
    ).toBe('No API key configured')
    expect((await generateProviderImage('openai', config, '', { prompt: 'x' })).error).toBe(
      'No image model configured',
    )
    expect(
      (await generateProviderImage('openai', config, 'gpt-image-2', { prompt: '  ' })).error,
    ).toBe('The prompt is empty')
    expect(fetchMock).not.toHaveBeenCalled()
  })
})
