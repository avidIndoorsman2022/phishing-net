import * as React from 'react';
import { Separator } from 'office-ui-fabric-react/lib/Separator';

export interface ISeparatorProp {
    content: string;
  }
  
export default class HorizontalSeparator extends React.Component<ISeparatorProp, {}> {
  constructor(props, context) {
    super(props, context);
  }
      
  public render(): JSX.Element {
    const content = this.props.content;

    return (
       <Separator>{content}</Separator>
    );
  }
}
