export type FileMetadata = {
    id: string;
    filename: string;
    size: number;
    content_type: string | null;
    last_modified: string | null;
};

export type ListFilesResponse = {
    files: string[];
};

export type ThumbnailResponse = {
    slide_index: number;
    image_base64: string;
};

export type SlideCountResponse = {
    file_id: string;
    slide_count: number;
};

export type AllThumbnailsResponse = {
    file_id: string;
    thumbnails: string[];
    format?: string;
};

export type PreviewParams = {
    tags: string[];
    sidebar_width: number;
    sidebar_item_height: number;
    sidebar_color_hex?: string;
    indicator_color_hex?: string;
    sidebar_item_font_color_hex?: string;
    sidebar_transparency?: number;
    sidebar_init_font_size?: number;
    vertically_center?: boolean;
    rounded_indicator?: boolean;
    center_text?: boolean;
    compact_indicator?: boolean;
};

export type PreviewThumbnailsResult = {
    thumbnails: string[];
    format: string;
};

export type PreviewProgressEvent = {
    stage: 'processing' | 'converting' | 'rendering' | 'done' | 'error';
    progress: number;
    message: string;
    current_slide?: number;
    total_slides?: number;
    thumbnails?: string[];
    format?: string;
};

export type ProcessParams = {
    tags: string[];
    sidebar_width: number;
    sidebar_item_height: number;
    transition_duration: number;
    apply_morph_transition: boolean;
    sidebar_color_hex?: string;
    indicator_color_hex?: string;
    sidebar_item_font_color_hex?: string;
    sidebar_transparency?: number;
    sidebar_init_font_size?: number;
    vertically_center?: boolean;
    rounded_indicator?: boolean;
    center_text?: boolean;
    compact_indicator?: boolean;
};

export type ProgressEvent = {
    stage: string;
    progress: number;
    message: string;
    job_id?: string;
    file_id?: string;
};

