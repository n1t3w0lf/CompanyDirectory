import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';

export interface IPropertyPaneSyncButtonProps {
  key: string;
  text: string;
  disabled?: boolean;
  description?: string;
  onClick: () => void;
}

export interface IPropertyPaneSyncButtonInternalProps extends IPropertyPaneSyncButtonProps {
  targetProperty: string;
  onRender: (elem: HTMLElement) => void;
  onDispose: (elem: HTMLElement) => void;
}

class PropertyPaneSyncButtonBuilder implements IPropertyPaneField<IPropertyPaneSyncButtonInternalProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneSyncButtonInternalProps;

  constructor(targetProperty: string, properties: IPropertyPaneSyncButtonProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.key,
      text: properties.text,
      disabled: properties.disabled,
      description: properties.description,
      onClick: properties.onClick,
      targetProperty: targetProperty,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    };
  }

  private onRender(elem: HTMLElement): void {
    const buttonElement = React.createElement(SyncButtonComponent, {
      text: this.properties.text,
      disabled: this.properties.disabled,
      description: this.properties.description,
      onClick: this.properties.onClick
    });
    ReactDom.render(buttonElement, elem);
  }

  private onDispose(elem: HTMLElement): void {
    ReactDom.unmountComponentAtNode(elem);
  }
}

interface ISyncButtonComponentProps {
  text: string;
  disabled?: boolean;
  description?: string;
  onClick: () => void;
}

interface ISyncButtonComponentState {
  isProcessing: boolean;
}

class SyncButtonComponent extends React.Component<ISyncButtonComponentProps, ISyncButtonComponentState> {
  constructor(props: ISyncButtonComponentProps) {
    super(props);
    this.state = {
      isProcessing: false
    };
  }

  private handleClick = async (): Promise<void> => {
    this.setState({ isProcessing: true });
    try {
      await this.props.onClick();
    } finally {
      this.setState({ isProcessing: false });
    }
  }

  public render(): JSX.Element {
    const buttonStyle: React.CSSProperties = {
      backgroundColor: '#0078d4',
      color: '#ffffff',
      border: 'none',
      padding: '8px 16px',
      fontSize: '14px',
      fontWeight: 600,
      cursor: this.props.disabled || this.state.isProcessing ? 'not-allowed' : 'pointer',
      borderRadius: '2px',
      opacity: this.props.disabled || this.state.isProcessing ? 0.6 : 1,
      width: '100%',
      marginTop: '4px'
    };

    const containerStyle: React.CSSProperties = {
      marginBottom: '8px'
    };

    const descriptionStyle: React.CSSProperties = {
      fontSize: '12px',
      color: '#605e5c',
      marginTop: '4px',
      fontStyle: 'italic'
    };

    return React.createElement('div', { style: containerStyle },
      React.createElement('button', {
        style: buttonStyle,
        onClick: this.handleClick,
        disabled: this.props.disabled || this.state.isProcessing
      }, this.state.isProcessing ? 'Syncing...' : this.props.text),
      this.props.description && React.createElement('div', { style: descriptionStyle }, this.props.description)
    );
  }
}

export function PropertyPaneSyncButton(targetProperty: string, properties: IPropertyPaneSyncButtonProps): IPropertyPaneField<IPropertyPaneSyncButtonInternalProps> {
  return new PropertyPaneSyncButtonBuilder(targetProperty, properties);
}
