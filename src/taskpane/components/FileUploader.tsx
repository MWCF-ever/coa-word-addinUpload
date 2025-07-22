// import React, { useRef, useState } from 'react';
// import { Stack, Text, PrimaryButton, DefaultButton, MessageBar, MessageBarType, Icon } from '@fluentui/react';

// interface FileUploaderProps {
//     onUpload: (file: File) => void;
//     disabled?: boolean;
// }

// export const FileUploader: React.FC<FileUploaderProps> = ({ onUpload, disabled = false }) => {
//     const fileInputRef = useRef<HTMLInputElement>(null);
//     const [selectedFile, setSelectedFile] = useState<File | null>(null);
//     const [error, setError] = useState<string | null>(null);

//     const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
//         const file = event.target.files?.[0];
//         if (file) {
//             if (file.type !== 'application/pdf') {
//                 setError('Please select a PDF file');
//                 setSelectedFile(null);
//                 return;
//             }
//             if (file.size > 10 * 1024 * 1024) { // 10MB limit
//                 setError('File size must be less than 10MB');
//                 setSelectedFile(null);
//                 return;
//             }
//             setError(null);
//             setSelectedFile(file);
//         }
//     };

//     const handleUploadClick = () => {
//         if (fileInputRef.current) {
//             fileInputRef.current.click();
//         }
//     };

//     const handleSubmit = () => {
//         if (selectedFile) {
//             onUpload(selectedFile);
//             setSelectedFile(null);
//             if (fileInputRef.current) {
//                 fileInputRef.current.value = '';
//             }
//         }
//     };

//     const handleClear = () => {
//         setSelectedFile(null);
//         setError(null);
//         if (fileInputRef.current) {
//             fileInputRef.current.value = '';
//         }
//     };

//     return (
//         <Stack tokens={{ childrenGap: 10 }}>
//             <Text variant="medium" className="field-label">Upload COA Document</Text>
            
//             <input
//                 ref={fileInputRef}
//                 type="file"
//                 accept=".pdf"
//                 onChange={handleFileSelect}
//                 style={{ display: 'none' }}
//             />

//             {error && (
//                 <MessageBar 
//                     messageBarType={MessageBarType.error}
//                     onDismiss={() => setError(null)}
//                 >
//                     {error}
//                 </MessageBar>
//             )}

//             {!selectedFile ? (
//                 <PrimaryButton
//                     text="Select PDF File"
//                     iconProps={{ iconName: 'Upload' }}
//                     onClick={handleUploadClick}
//                     disabled={disabled}
//                     styles={{ root: { width: '100%' } }}
//                 />
//             ) : (
//                 <Stack tokens={{ childrenGap: 10 }}>
//                     <Stack 
//                         horizontal 
//                         verticalAlign="center" 
//                         tokens={{ childrenGap: 10 }}
//                         className="file-info"
//                     >
//                         <Icon iconName="PDF" style={{ fontSize: 20, color: '#D13438' }} />
//                         <Text>{selectedFile.name}</Text>
//                         <Text variant="small">({(selectedFile.size / 1024).toFixed(2)} KB)</Text>
//                     </Stack>
                    
//                     <Stack horizontal tokens={{ childrenGap: 10 }}>
//                         <PrimaryButton
//                             text="Process File"
//                             iconProps={{ iconName: 'Processing' }}
//                             onClick={handleSubmit}
//                             disabled={disabled}
//                             styles={{ root: { flex: 1 } }}
//                         />
//                         <DefaultButton
//                             text="Clear"
//                             iconProps={{ iconName: 'Clear' }}
//                             onClick={handleClear}
//                             disabled={disabled}
//                             styles={{ root: { flex: 1 } }}
//                         />
//                     </Stack>
//                 </Stack>
//             )}
//         </Stack>
//     );
// };