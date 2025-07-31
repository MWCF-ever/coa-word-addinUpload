import React, { useMemo } from 'react';
import { Dropdown, IDropdownOption, Stack, Text } from '@fluentui/react';
import { Compound, Template } from '../../types';

// 使用React.memo优化CompoundSelector
interface CompoundSelectorProps {
    compounds: Compound[];
    selectedCompound?: Compound;
    onSelect: (compound: Compound) => void;
    disabled?: boolean;
}

export const CompoundSelector = React.memo<CompoundSelectorProps>(({
    compounds,
    selectedCompound,
    onSelect,
    disabled = false
}) => {
    // 使用useMemo缓存options
    const options = useMemo<IDropdownOption[]>(() => 
        compounds.map(compound => ({
            key: compound.id,
            text: `${compound.code} - ${compound.name}`,
            data: compound
        })),
        [compounds]
    );

    const handleChange = React.useCallback((
        event: React.FormEvent<HTMLDivElement>, 
        option?: IDropdownOption
    ) => {
        if (option?.data) {
            onSelect(option.data as Compound);
        }
    }, [onSelect]);

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
}, (prevProps, nextProps) => {
    // 自定义比较函数，只在必要时重新渲染
    return (
        prevProps.compounds === nextProps.compounds &&
        prevProps.selectedCompound?.id === nextProps.selectedCompound?.id &&
        prevProps.disabled === nextProps.disabled &&
        prevProps.onSelect === nextProps.onSelect
    );
});

// 使用React.memo优化TemplateSelector
interface TemplateSelectorProps {
    templates: Template[];
    selectedTemplate?: Template;
    onSelect: (template: Template) => void;
    disabled?: boolean;
}

export const TemplateSelector = React.memo<TemplateSelectorProps>(({
    templates,
    selectedTemplate,
    onSelect,
    disabled = false
}) => {
    // 使用useMemo缓存options
    const options = useMemo<IDropdownOption[]>(() => 
        templates.map(template => ({
            key: template.id,
            text: `${template.region} Template`,
            data: template
        })),
        [templates]
    );

    const handleChange = React.useCallback((
        event: React.FormEvent<HTMLDivElement>, 
        option?: IDropdownOption
    ) => {
        if (option?.data) {
            onSelect(option.data as Template);
        }
    }, [onSelect]);

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
}, (prevProps, nextProps) => {
    // 自定义比较函数
    return (
        prevProps.templates === nextProps.templates &&
        prevProps.selectedTemplate?.id === nextProps.selectedTemplate?.id &&
        prevProps.disabled === nextProps.disabled &&
        prevProps.onSelect === nextProps.onSelect
    );
});