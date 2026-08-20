import { z } from 'zod';
import zodToJsonSchema from 'zod-to-json-schema';

const propertiesSchema = z.object({
  message: z.string().describe('A message to display.'),
  dataSource: z
    .enum(['mock', 'graph'])
    .optional()
    .describe(
      'Where My Day reads its data from: bundled demo data (mock, the default) or live Microsoft Graph data (graph).'
    ),
  newsSiteUrl: z
    .string()
    .optional()
    .describe(
      'Absolute URL of the SharePoint site to read news posts from in graph mode. Defaults to the tenant root site.'
    )
});

export type IMyDayCopilotComponentProperties = z.infer<typeof propertiesSchema>;

export default zodToJsonSchema(propertiesSchema);
