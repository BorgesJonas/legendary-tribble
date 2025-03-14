export class PowerBiReportDetails {
  constructor(
    public reportId: string,
    public reportName: string,
    public embedUrl: string,
  ) {}
}

export class EmbedConfig {
  constructor(
    public type?: string,
    public reportsDetail?: PowerBiReportDetails[],
    public embedToken?: any,
  ) {}
}
