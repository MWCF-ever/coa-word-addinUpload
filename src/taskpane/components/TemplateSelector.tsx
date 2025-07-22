import React from 'react';
import { Dropdown, IDropdownOption, Stack, Text } from '@fluentui/react';
import { Template } from '../../types';

interface TemplateSelectorProps {
    templates: Template[];
    selectedTemplate?: Template;
    onSelect: (template: Template) => void;
    disabled?: boolean;
}

export const TemplateSelector: React.FC<TemplateSelectorProps> = ({
    templates,
    selectedTemplate,
    onSelect,
    disabled = false
}) => {
    const options: IDropdownOption[] = templates.map(template => ({
        key: template.id,
        text: `${template.region} Template`,
        data: template
    }));

    const handleChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option && option.data) {
            onSelect(option.data as Template);
        }
    };

    return (
        <Stack tokens={{ childrenGap: 5 }}>
            <Text variant="medium" className="field-label">Select Template Region</Text>
            <Dropdown
                placeholder="Choose a template"
                options={options}
                selectedKey={selectedTemplate?.id}
                onChange={handleChange}
                disabled={disabled || templates.length === 0}
                styles={{
                    dropdown: { width: '100%' }
                }}
            />
        </Stack>
    );
};