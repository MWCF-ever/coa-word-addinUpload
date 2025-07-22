import React from 'react';
import { Dropdown, IDropdownOption, Stack, Text } from '@fluentui/react';
import { Compound } from '../../types';

interface CompoundSelectorProps {
    compounds: Compound[];
    selectedCompound?: Compound;
    onSelect: (compound: Compound) => void;
    disabled?: boolean;
}

export const CompoundSelector: React.FC<CompoundSelectorProps> = ({
    compounds,
    selectedCompound,
    onSelect,
    disabled = false
}) => {
    const options: IDropdownOption[] = compounds.map(compound => ({
        key: compound.id,
        text: `${compound.code} - ${compound.name}`,
        data: compound
    }));

    const handleChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option && option.data) {
            onSelect(option.data as Compound);
        }
    };

    return (
        <Stack tokens={{ childrenGap: 5 }}>
            <Text variant="medium" className="field-label">Select Compound</Text>
            <Dropdown
                placeholder="Choose a compound"
                options={options}
                selectedKey={selectedCompound?.id}
                onChange={handleChange}
                disabled={disabled}
                styles={{
                    dropdown: { width: '100%' }
                }}
            />
        </Stack>
    );
};