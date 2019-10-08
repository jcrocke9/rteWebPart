import * as React from 'react';
import styles from './RichTextEditor.module.scss';

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { sp } from "@pnp/sp";

export interface IRichTextEditorProps {
  description: string;
}

export interface IRichTextEditorState {
  value: string;
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
      value: ""
    }
    this.onChange = this.onChange.bind(this);
    this.onSubmit = this.onSubmit.bind(this);
  }

  public render(): React.ReactElement<IRichTextEditorProps> {
    return (
      <div className={styles.richTextEditor} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <RichText
                value={this.state.value}
                onChange={this.onChange}
                className={styles.aopRte}
                isEditMode={true}
              />
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
        //this.onChange(item.ProjectOverviewVal);
        this.setState({ value: item.ProjectOverviewVal });
      });
  }
  private onChange = (newText: string): string => {
    this.setState({ value: newText });
    console.log(newText);
    return newText;
  }
  private onSubmit = (e: any): void => {
    sp.web.lists.getByTitle("AOP Projects2").items.getById(1).update({
      ProjectOverviewVal: this.state.value,
    })
  }
}