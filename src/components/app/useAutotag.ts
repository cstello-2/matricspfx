import * as React from 'react';
import { DataContext } from '../../core/datasources';
import { AutoTagger, AutotagAccess, TAGGABLE_KINDS, checkAutotagAccess } from '../../core/ai/autotag';
import { OpenAiCompatibleClient } from '../../core/ai/openai';
import { ListModel } from './useListModel';

export interface AutotagState {
  tagger?: AutoTagger;
  access?: AutotagAccess;
  error?: string;
}

export function useAutotag(model: ListModel | undefined, data: DataContext): AutotagState {
  const config = model && model.config.autotag;
  const field = config && model ? model.fieldMap.get(config.field) : undefined;
  const service = model ? model.service : undefined;
  const fieldsLoaded = !!model && !!model.fields;
  const [access, setAccess] = React.useState<AutotagAccess | undefined>(undefined);

  const setup = React.useMemo((): { tagger?: AutoTagger; error?: string } => {
    if (!config || !service || !fieldsLoaded) return {};
    if (!field) return { error: `Auto-tagging: the list has no column named "${config.field}".` };
    if (TAGGABLE_KINDS.indexOf(field.kind) < 0) return { error: `Auto-tagging: "${field.title}" is a ${field.kind} column; use a choice, lookup or text column.` };
    if (!config.api || !config.api.baseUrl || !config.api.embeddingModel) return { error: 'Auto-tagging needs api.baseUrl and api.embeddingModel.' };
    const client = new OpenAiCompatibleClient(data.http, config.api, data.webUrl);
    return { tagger: new AutoTagger(service, client, config, field) };
  }, [config, service, field, fieldsLoaded, data]);

  React.useEffect(() => {
    if (!config || !service || !setup.tagger) return;
    let live = true;
    setAccess(undefined);
    checkAutotagAccess(service, config, data.user).then(
      result => {
        if (live) setAccess(result);
      },
      (e: Error) => {
        if (live) setAccess({ allowed: false, reason: `Could not check permissions: ${e.message}` });
      },
    );
    return () => {
      live = false;
    };
  }, [config, service, setup.tagger, data.user]);

  return { tagger: setup.tagger, access, error: setup.error };
}
