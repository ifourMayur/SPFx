import { IProjectTemplateFolder } from './ProjectTemplateFolder';
import { planFolderTree, sanitizeFolderName } from './SharePointFolder';

/** A template folder with one named sub-folder chain, built from the level names given. */
function folder(folderName: string, ...subFolderNames: string[]): IProjectTemplateFolder {
  return {
    foldersId: 5,
    folderName,
    lstProjectsubFolderViewModel: subFolderNames.length
      ? [
          {
            id: 11,
            folderID: 5,
            subFolderName: subFolderNames[0],
            lstProjectsubsubFolderViewModel: subFolderNames.slice(1).reduceRight(
              (children: { id: number; folderID: number; subFolderName: string; lstProjectsubsubFolderViewModel: never[] }[], name: string) => [
                {
                  id: 0,
                  folderID: 5,
                  subFolderName: name,
                  lstProjectsubsubFolderViewModel: children as never[]
                }
              ],
              []
            )
          }
        ]
      : []
  };
}

describe('sanitizeFolderName', () => {
  it('keeps a name SharePoint already accepts', () => {
    expect(sanitizeFolderName('Drawings v1.2')).toBe('Drawings v1.2');
  });

  it('replaces every character SharePoint refuses with a space', () => {
    expect(sanitizeFolderName('Plans<Old>')).toBe('Plans Old');
    expect(sanitizeFolderName('A/B\\C:D*E?F"G|H<I>J')).toBe('A B C D E F G H I J');
  });

  it('leaves # and % alone - SharePoint Online allows them now', () => {
    expect(sanitizeFolderName('Phase #2 (50% done)')).toBe('Phase #2 (50% done)');
  });

  it('strips the reserved _vti_ sequence', () => {
    expect(sanitizeFolderName('report_vti_files')).toBe('report files');
  });

  it('trims a name that may not begin or end with a period', () => {
    expect(sanitizeFolderName('.Drawings.')).toBe('Drawings');
    expect(sanitizeFolderName('  . Drawings .  ')).toBe('Drawings');
  });

  it('returns an empty string when nothing usable is left', () => {
    expect(sanitizeFolderName('///')).toBe('');
    expect(sanitizeFolderName('   ')).toBe('');
    expect(sanitizeFolderName('')).toBe('');
  });
});

describe('planFolderTree', () => {
  it('puts the project root first', () => {
    const paths: string[] = planFolderTree('Roof replacement', [folder('Drawings')]);

    expect(paths[0]).toBe('Roof replacement');
  });

  it('walks all four levels', () => {
    const paths: string[] = planFolderTree('Roof replacement', [
      folder('Drawings', 'Architectural', 'Floor plans', 'Ground floor')
    ]);

    expect(paths).toEqual([
      'Roof replacement',
      'Roof replacement/Drawings',
      'Roof replacement/Drawings/Architectural',
      'Roof replacement/Drawings/Architectural/Floor plans',
      'Roof replacement/Drawings/Architectural/Floor plans/Ground floor'
    ]);
  });

  it('lists every parent before its children, which is the order they must be created in', () => {
    const paths: string[] = planFolderTree('P', [
      folder('A', 'A1'),
      folder('B', 'B1')
    ]);

    paths.forEach((path: string, index: number): void => {
      const lastSlash: number = path.lastIndexOf('/');

      if (lastSlash > 0) {
        expect(paths.indexOf(path.substring(0, lastSlash))).toBeLessThan(index);
      }
    });
  });

  it('sanitizes the project name into the root folder', () => {
    const paths: string[] = planFolderTree('Roof/replacement', [folder('Drawings')]);

    expect(paths).toEqual(['Roof replacement', 'Roof replacement/Drawings']);
  });

  it('returns nothing when the project name sanitizes away - there is no root to build on', () => {
    expect(planFolderTree('///', [folder('Drawings')])).toEqual([]);
  });

  it('skips a nameless folder and its children rather than inventing a name', () => {
    const paths: string[] = planFolderTree('P', [folder('', 'Orphan'), folder('Contracts')]);

    expect(paths).toEqual(['P', 'P/Contracts']);
  });

  it('plans a repeated path once, but still hangs its children off it', () => {
    const paths: string[] = planFolderTree('P', [
      folder('Drawings', 'A'),
      // Sanitizes to the same name, so both would be one folder in the library.
      folder('Drawings*', 'B')
    ]);

    expect(paths).toEqual(['P', 'P/Drawings', 'P/Drawings/A', 'P/Drawings/B']);
  });

  it('treats names differing only in case as one folder, as SharePoint does', () => {
    const paths: string[] = planFolderTree('P', [folder('Drawings'), folder('DRAWINGS')]);

    expect(paths).toEqual(['P', 'P/Drawings']);
  });

  it('plans just the root when the template defines no folders', () => {
    expect(planFolderTree('P', [])).toEqual(['P']);
    expect(planFolderTree('P', undefined)).toEqual(['P']);
  });

  it('treats the synthetic BIM folder as any other, including its Model child', () => {
    // What the API appends when isBIM is set: both nodes carry id 0.
    const paths: string[] = planFolderTree('P', [
      { foldersId: 0, folderName: 'BIM', lstProjectsubFolderViewModel: [{ id: 0, folderID: 0, subFolderName: 'Model' }] }
    ]);

    expect(paths).toEqual(['P', 'P/BIM', 'P/BIM/Model']);
  });
});
