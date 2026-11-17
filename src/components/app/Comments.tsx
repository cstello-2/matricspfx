import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SpComment, SpListService } from '../../core/sharepoint/listService';
import { AppStyles } from './appStyles';
import { relativeTime } from './time';

export interface CommentsProps {
  service: SpListService;
  itemId: number;
  currentUser?: string;
  styles: AppStyles;
  locale?: string;
}

export function Comments({ service, itemId, currentUser, styles, locale }: CommentsProps): React.ReactElement {
  const [comments, setComments] = React.useState<SpComment[] | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [draft, setDraft] = React.useState('');
  const [posting, setPosting] = React.useState(false);

  React.useEffect(() => {
    let live = true;
    setComments(undefined);
    setError(undefined);
    service.comments(itemId).then(
      list => {
        if (live) setComments(list.slice().sort((a, b) => a.createdDate.localeCompare(b.createdDate)));
      },
      (e: Error) => {
        if (live) setError(e.message);
      },
    );
    return () => {
      live = false;
    };
  }, [service, itemId]);

  const post = (): void => {
    const text = draft.trim();
    if (!text || posting) return;
    setPosting(true);
    setError(undefined);
    service.addComment(itemId, text).then(
      comment => {
        setComments(list => (list || []).concat(comment));
        setDraft('');
        setPosting(false);
      },
      (e: Error) => {
        setError(e.message);
        setPosting(false);
      },
    );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {!comments && !error && <Spinner size={SpinnerSize.small} label="Loading comments" labelPosition="right" />}
      {comments && !comments.length && <div className={styles.meta}>No comments yet.</div>}
      {(comments || []).map(c => (
        <div key={c.id} className={styles.comment}>
          <Persona text={c.authorName} size={PersonaSize.size32} hidePersonaDetails />
          <div>
            <div className={styles.commentHead}>
              <span className={styles.commentAuthor}>{c.authorName}</span>
              <span className={styles.commentTime} title={new Date(c.createdDate).toLocaleString(locale)}>{relativeTime(c.createdDate, locale)}</span>
            </div>
            <div className={styles.commentText}>{c.text}</div>
          </div>
        </div>
      ))}
      <div className={styles.composer}>
        <Persona text={currentUser || 'You'} size={PersonaSize.size32} hidePersonaDetails />
        <div>
          <TextField
            multiline
            autoAdjustHeight
            rows={draft ? 3 : 1}
            placeholder="Add a comment…"
            value={draft}
            onChange={(_, v) => setDraft(v || '')}
            onKeyDown={e => {
              if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) post();
            }}
            ariaLabel="Add a comment"
            disabled={posting}
          />
          {draft && (
            <div className={styles.composerActions}>
              <PrimaryButton text={posting ? 'Saving…' : 'Save'} onClick={post} disabled={posting || !draft.trim()} />
              <DefaultButton text="Cancel" onClick={() => setDraft('')} disabled={posting} />
              <span className={styles.meta}>Ctrl+Enter to save</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
