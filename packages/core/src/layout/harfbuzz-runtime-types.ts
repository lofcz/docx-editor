export type HarfBuzzModule = typeof import('harfbuzzjs');
type HarfBuzzBlob = InstanceType<HarfBuzzModule['Blob']>;
export type HarfBuzzBuffer = InstanceType<HarfBuzzModule['Buffer']>;
type HarfBuzzFace = InstanceType<HarfBuzzModule['Face']>;
type HarfBuzzFont = InstanceType<HarfBuzzModule['Font']>;

export interface ActiveHarfBuzzFont {
  readonly identity: string;
  readonly blob: HarfBuzzBlob;
  readonly face: HarfBuzzFace;
  readonly font: HarfBuzzFont;
  readonly unitsPerEm: number;
  readonly byteLength: number;
}
