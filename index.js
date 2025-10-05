import 'dotenv/config';
import express from 'express';
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  MemoryStorage,
  ConversationState,
  CardFactory
} from 'botbuilder';

const {
  PORT = 3978,
  MicrosoftAppId,
  MicrosoftAppPassword,
  MicrosoftAppType = 'SingleTenant',   // 'SingleTenant' o 'MultiTenant'
  MicrosoftAppTenantId,                // requerido si SingleTenant
  BACKEND_URL                          // ej: https://TU-BACKEND.azurewebsites.net
} = process.env;

if (!MicrosoftAppId || !MicrosoftAppPassword || !BACKEND_URL) {
  console.error('Faltan vars: MicrosoftAppId, MicrosoftAppPassword, BACKEND_URL');
  process.exit(1);
}

const auth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId,
  MicrosoftAppPassword,
  MicrosoftAppType,
  MicrosoftAppTenantId
});
const adapter = new CloudAdapter(auth);
const conversationState = new ConversationState(new MemoryStorage());

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]:', error);
  await context.sendActivity('Lo siento, ocurrió un error procesando tu mensaje.');
};

const app = express();
app.use(express.json());

function buildAnswerCard({ title = 'ARCO Buddy', answer, files = [] }) {
  const maxToShow = 5;
  const shown = files.slice(0, maxToShow);
  const sourcesMarkdown = shown.length ? shown.map(f => `• [${f.name}](${f.webUrl})`).join('\n') : '';
  const actions = shown.slice(0, 3).map(f => ({
    type: 'Action.OpenUrl',
    title: `Abrir: ${f.name}`.slice(0, 40),
    url: f.webUrl
  }));
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: title, size: 'Medium', weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: answer || 'No encontré información suficiente en los documentos.', wrap: true },
      ...(sourcesMarkdown ? [
        { type: 'TextBlock', text: 'Fuentes', weight: 'Bolder', spacing: 'Medium' },
        { type: 'TextBlock', text: sourcesMarkdown, wrap: true }
      ] : [])
    ],
    actions
  };
}

async function onMessage(context) {
  const text = (context.activity.text || '').trim();
  if (!text) { await context.sendActivity('¿Puedes escribir tu pregunta?'); return; }
  await context.sendActivities([{ type: 'typing' }]);

  let result;
  try {
    const resp = await fetch(`${BACKEND_URL}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query: text }),
      signal: AbortSignal.timeout(25000)
    });
    result = await resp.json();
    if (!resp.ok) throw new Error(result?.error || `HTTP ${resp.status}`);
  } catch (e) {
    console.error('Error llamando BACKEND /chat:', e);
    await context.sendActivity('No pude consultar el conocimiento interno. Intenta de nuevo en un momento.');
    return;
  }

  // DESPUÉS de llamar al backend:
  const answer = (result?.answer || '').toString().trim();
  const files  = Array.isArray(result?.topFiles) ? result.topFiles : [];

  // (debug temporal)
  console.log('[bot] answerPreview:', answer.slice(0, 160));

  const card = CardFactory.adaptiveCard(buildAnswerCard({
    title: 'ARCO Buddy',
    answer,    // <— usar la respuesta del backend
    files
  }));
await context.sendActivity({ attachments: [card] });
}

app.post('/api/messages', (req, res) => {
  adapter.process(req, res, async (context) => {
    if (context.activity.type === 'message') {
      await onMessage(context);
    } else {
      await context.sendActivity(`Evento: ${context.activity.type}`);
    }
    await conversationState.saveChanges(context, false);
  });
});

app.get('/', (_req, res) => res.send('Bot OK'));

app.listen(PORT, () => console.log(`Arco Buddy listening on http://localhost:${PORT}`));
