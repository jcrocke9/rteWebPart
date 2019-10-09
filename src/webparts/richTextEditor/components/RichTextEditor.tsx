import * as React from 'react';
import styles from './RichTextEditor.module.scss';

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { sp } from "@pnp/sp";

export interface IRichTextEditorProps {
  description: string;
}

export interface IRichTextEditorState {
  value?: string;
  nowLoad?: boolean;
}

export interface IAopRichText {
  ID?: number;
  ProjectObjectiveVal?: string;
  ProjectOverviewVal?: string;
}

export default class RichTextEditor extends React.Component<IRichTextEditorProps, IRichTextEditorState> {
  constructor(props) {
    super(props);
    this.state = {
      value: "",
      nowLoad: false
    }
    this.onChange = this.onChange.bind(this);
    this.onSubmit = this.onSubmit.bind(this);
  }

  public render(): React.ReactElement<IRichTextEditorProps> {
    const { value, nowLoad } = this.state;
    return (
      <div className={styles.richTextEditor} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {nowLoad &&
                <RichText
                  value={value}
                  onChange={this.onChange}
                  className={styles.aopRte}
                  isEditMode={true}
                />
              }
              <p>
                <a onClick={(e) => this.onSubmit(e)}><PrimaryButton text="Submit" /></a>
              </p>
            </div>
          </div>
        </div>
      </div >
    );
  }
  private componentDidMount(): void {
    sp.web.lists.getByTitle("AOP Projects2").items.getById(1).get()
      .then((item: IAopRichText) => {
        this.setState({ value: item.ProjectOverviewVal });
      })
      .then(() => {
        this.setState({nowLoad: true})
      })
  }
  private onChange = (newText: string): string => {
    this.setState({ value: newText });
    return newText;
  }
  private onSubmit = (e: any): void => {
    sp.web.lists.getByTitle("AOP Projects2").items.getById(1).update({
      ProjectOverviewVal: this.state.value,
    })
  }
}