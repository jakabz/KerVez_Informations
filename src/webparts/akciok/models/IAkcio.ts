export interface IAkcio {
    title: string;
    actionType: string;
    status: string;
    startDate: Date | undefined;
    endDate: Date | undefined;
    publishStartDate: Date | undefined;
    description: string;
    responsiblePersonLoginName: string;
    responsibleEmail: string;
    relatedDocuments: string;
    relatedShopsType: string;
    relatedShops: string;
    relatedProducts: string;
    images: string;
}

export interface IAkcioListItem {
    id: number;
    title: string;
    actionType: string;
    status: string;
    startDate: string | null;
    endDate: string | null;
    publishStartDate: string | null;
    responsiblePersonTitle: string;
    relatedProducts: string;
}

export interface IAkcioDetails extends IAkcio {
    id: number;
    responsiblePersonTitle?: string;
    responsiblePersonEmail?: string;
}

export const defaultAkcio: IAkcio = {
    title: '',
    actionType: '',
    status: 'Piszkozat',
    startDate: undefined,
    endDate: undefined,
    publishStartDate: undefined,
    description: '',
    responsiblePersonLoginName: '',
    responsibleEmail: '',
    relatedDocuments: '',
    relatedShopsType: 'all',
    relatedShops: '',
    relatedProducts: '',
    images: '',
};
