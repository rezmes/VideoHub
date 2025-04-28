// src/shared/components/GenericDropdown.tsx
import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export interface IGenericDropdownProps {
  label: string;
  options: IDropdownOption[];
  selectedKey: string | number | null;
  onChanged: (option?: IDropdownOption) => void;
  placeHolder?: string;
  disabled?: boolean;
  className?: string;
}

export class GenericDropdown extends React.Component<IGenericDropdownProps, {}> {
  public render(): React.ReactElement<IGenericDropdownProps> {
    const { label, options, selectedKey, onChanged, placeHolder, disabled, className } = this.props;
    console.log(`Rendering dropdown '${label}' with selectedKey:`, selectedKey);
    console.log(`Options for '${label}':`, options);
    
    return (
      <Dropdown
        label={label}
        options={options}
        selectedKey={selectedKey}
        onChanged={onChanged}
        placeHolder={placeHolder}
        disabled={disabled}
        className={className}
      />
    );
  }
}

export default GenericDropdown;
