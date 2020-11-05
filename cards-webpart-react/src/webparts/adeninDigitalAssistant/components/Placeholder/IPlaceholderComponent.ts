export interface IPlaceholderProps {
    description: string;
    iconName: string;
    iconText: string;
    configButtonLabel?: string;
    assistantButtonLabel?: string;
    contentClassName?: string;
    apiURL?: string;
    embedType?: string;
    cardId?: string;
    onConfigure?: () => void;
}