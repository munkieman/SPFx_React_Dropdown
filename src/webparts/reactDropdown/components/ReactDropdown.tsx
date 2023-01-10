import * as React from 'react';
import styles from './ReactDropdown.module.scss';
import { IReactDropdownProps } from './IReactDropdownProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ComboBox, IComboBoxOption, IComboBox, PrimaryButton } from 'office-ui-fabric-react/lib/index';
import { getGUID } from "@pnp/common";
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IStates {
  Select1: any;
  Select2: any;
}

export default class ReactDropdown extends React.Component<IReactDropdownProps, IStates> {
  constructor(props: IReactDropdownProps | Readonly<IReactDropdownProps>) {
    super(props);
    this.state = {
      Select1: "",
      Select2: ""
    };
  }

  private async Save() {
    let web = Web(this.props.webURL);
    //alert('saving '+this.state.medicalSelect);
    //sp.web.lists.getByTitle("Audit Tool Data").items.add({
    //  Medical: this.state.medicalSelect
    //});
    await web.lists.getByTitle("Combobox_Test").items.add({
      Title: getGUID(),
      Choice1: this.state.Select1,
      Choice2: this.state.Select2
    });
    //.then(i => {
    //  console.log(i);
    //});
    alert("Submitted Successfully");
  }

  public onChange1 = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({ Select1: option.key });
    
  }

  public onChange2 = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({ Select2: option.key });
  }  

  public render(): React.ReactElement<IReactDropdownProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.reactDropdown} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
        <div>
          <h1>React JS ComboBox Examples</h1>
          <div>
            <ComboBox
              placeholder="Please Choose"
              selectedKey={this.state.Select1}
              label="Medicals"
              autoComplete="on"
              options={this.props.Choice1}
              onChange={this.onChange1}
            />
          </div>
          {this.state.Select1=="Yes" ?  
            <div>choice was yes</div>
            : null
          }          
          <div>
            <ComboBox className="hidden"
              placeholder="Please Choose"
              selectedKey={this.state.Select2}
              label="Assessment"
              autoComplete="on"
              options={this.props.Choice2}
              onChange={this.onChange2}
            />
          </div>
          <div>
            <br />
            <br />
            <PrimaryButton onClick={() => this.Save()}>Submit</PrimaryButton>
          </div>
        </div>
      </section>
    );
  }
}
