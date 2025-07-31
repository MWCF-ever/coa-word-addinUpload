import React, { useMemo } from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, SelectionMode, Stack, Text } from '@fluentui/react';
import { BatchData } from '../../types';

interface BatchDataTableProps {
    batchDataList: BatchData[];
}

export const BatchDataTable = React.memo<BatchDataTableProps>(({ batchDataList }) => {
    // 使用useMemo缓存列定义
    const columns = useMemo<IColumn[]>(() => [
        {
            key: 'batch_number',
            name: 'Batch Number',
            fieldName: 'batch_number',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
        },
        {
            key: 'manufacture_date',
            name: 'Manufacture Date',
            fieldName: 'manufacture_date',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
        },
        {
            key: 'manufacturer',
            name: 'Manufacturer',
            fieldName: 'manufacturer',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            onRender: (item: BatchData) => {
                const short = item.manufacturer.includes('Changzhou SynTheAll') 
                    ? 'Changzhou STA' 
                    : item.manufacturer;
                return <span title={item.manufacturer}>{short}</span>;
            }
        },
        {
            key: 'test_count',
            name: 'Test Results',
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            onRender: (item: BatchData) => {
                const count = Object.values(item.test_results || {})
                    .filter(v => v && v !== 'TBD' && v !== 'ND' && v !== '')
                    .length;
                const total = Object.keys(item.test_results || {}).length;
                return <Text>{`${count}/${total}`}</Text>;
            }
        }
    ], []);

    // 使用useMemo缓存处理后的数据
    const items = useMemo(() => 
        batchDataList.map((batch, index) => ({
            ...batch,
            key: batch.batch_number || `batch-${index}`
        })),
        [batchDataList]
    );

    if (batchDataList.length === 0) {
        return null;
    }

    return (
        <Stack tokens={{ childrenGap: 10 }}>
            <Text variant="medium" block>
                Batch Data Summary ({batchDataList.length} batches)
            </Text>
            <DetailsList
                items={items}
                columns={columns}
                setKey="batchDataList"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                isHeaderVisible={true}
                styles={{
                    root: {
                        maxHeight: '300px',
                        overflowY: 'auto'
                    }
                }}
            />
        </Stack>
    );
});