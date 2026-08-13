export interface NamedDroppedFile {
  name: string;
}

export function isPresentationStudioProjectName(name: string): boolean {
  return /\.pstudio(?:-secure)?$/i.test(name.trim());
}

export function projectPackageFromDrop<T extends NamedDroppedFile>(files: readonly T[]): T | undefined {
  const projects = files.filter((file) => isPresentationStudioProjectName(file.name));
  if (projects.length === 0) return undefined;
  if (projects.length !== 1 || files.length !== 1) throw new Error("Drop one .pstudio project at a time without additional files.");
  return projects[0];
}
