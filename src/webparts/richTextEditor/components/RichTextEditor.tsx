import * as React from 'react';
import styles from './RichTextEditor.module.scss';
import '@pnp/polyfill-ie11';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { sp } from '@pnp/sp';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IRichTextEditorProps {
  description: string;
}

export interface IRichTextEditorState {
  aopItemId?: number;
  field?: string;
  value?: string;
  nowLoad?: boolean;
  savedNote?: boolean;
  saving?: boolean;
}

export interface IAopRichText {
  ID?: number;
  ProjectObjectiveVal?: string;
  ProjectOverviewVal?: string;
}

export default class RichTextEditor extends React.Component<IRichTextEditorProps, IRichTextEditorState> {
  constructor(props: IRichTextEditorProps) {
    super(props);
    this.state = {
      aopItemId: -1,
      field: '',
      value: '',
      nowLoad: false,
      savedNote: false,
      saving: false
    };
    this.onChange = this.onChange.bind(this);
    this.onSubmit = this.onSubmit.bind(this);
  }

  public render(): React.ReactElement<IRichTextEditorProps> {
    const { value, nowLoad, savedNote, saving } = this.state;
    return (
      <div className={styles.richTextEditor} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {nowLoad ?
                <div>
                  <RichText
                    value={value}
                    onChange={this.onChange}
                    className={styles.aopRte}
                    isEditMode={true}
                  />
                  <div style={{ marginTop: '8px' }}>
                    {savedNote ?
                      <div style={{ color: '#333', backgroundColor: '#dff6dd' }}>
                        <span>
                          <Icon iconName='Completed' style={{ color: '#107c10', margin: '16px 0 16px 16px' }} />
                          &nbsp;Edits have been saved, please click the close button above.
                        </span>
                      </div>
                      :
                      saving ?
                        <div style={{ float: 'left', padding: '8px 16px' }}>
                          <Spinner size={SpinnerSize.small} />
                        </div>
                        :
                        <div>
                          <a onClick={(e) => this.onSubmit(e)}><DefaultButton text='Save' /></a>
                        </div>
                    }
                  </div>
                </div>
                :
                <Spinner size={SpinnerSize.large} />
              }
            </div>
          </div>
        </div>
      </div >
    );
  }
  private componentDidMount(): void {
    const urlParams: URLSearchParams = new URLSearchParams(window.location.search);
    const aopItemId: string = urlParams.get('aopItemId');
    const field: string = urlParams.get('OnlyIncludeOneField');
    sp.web.lists.getByTitle('AOP Projects2').items.getById(+aopItemId)
      .select('ProjectObjectiveVal', 'ProjectOverviewVal')
      .get()
      .then((item: IAopRichText) => {
        this.setState({
          aopItemId: +aopItemId,
          field: field,
          value: item[field]
        });
      })
      .then(() => {
        this.setState({ nowLoad: true });
      });
  }
  private onChange = (newText: string): string => {
    this.setState({ value: newText });
    return newText;
  }
  private onSubmit = (e: React.MouseEvent<HTMLAnchorElement>): void => {
    this.setState({ saving: true });
    sp.web.lists.getByTitle('AOP Projects2').items.getById(this.state.aopItemId).update({
      [this.state.field]: this.state.value
    })
      .then(() => {
        this.setState({ savedNote: true, saving: false });
      })
      .catch();

  }
}