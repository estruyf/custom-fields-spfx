declare interface ICustomFieldsStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'customFieldsStrings' {
  const strings: ICustomFieldsStrings;
  export = strings;
}
