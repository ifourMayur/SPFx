import { IBuildingForm, createEmptyBuildingForm } from './Building';
import { IProjectListRow, nextProjectId, upsertProjectRow } from './ProjectListRow';

function row(overrides?: Partial<IProjectListRow>): IProjectListRow {
  return {
    id: 1,
    name: 'Existing project',
    isFavorite: false,
    administratie: 'none',
    voorbereiding: 'none',
    ontwerp: 'none',
    uitvoering: 'none',
    revisie: 'none',
    ...overrides
  };
}

function form(overrides?: Partial<IBuildingForm>): IBuildingForm {
  return { ...createEmptyBuildingForm(), buildingName: 'Roof replacement', ...overrides };
}

describe('nextProjectId', () => {
  it('starts at 1 for an empty list', () => {
    expect(nextProjectId([])).toBe(1);
  });

  it('is one past the highest existing id, not the row count', () => {
    expect(nextProjectId([row({ id: 3 }), row({ id: 7 }), row({ id: 5 })])).toBe(8);
  });
});

describe('upsertProjectRow - creating', () => {
  // First, not last: the listing pages at 10 rows, so an appended project would land on
  // page 2 and be invisible on the page the user is returned to after saving.
  it('puts a newly created project first, where the user will see it', () => {
    const result = upsertProjectRow([row({ id: 1 })], form({ id: 0 }));

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0].name).toBe('Roof replacement');
  });

  it('keeps the projects that were already there', () => {
    const result = upsertProjectRow([row({ id: 1, name: 'Existing' })], form({ id: 0 }));

    expect(result.rows[1].name).toBe('Existing');
  });

  it('assigns the next free id and reports it', () => {
    const result = upsertProjectRow([row({ id: 4 })], form({ id: 0 }));

    expect(result.id).toBe(5);
    expect(result.rows[0].id).toBe(5);
  });

  it('gives a new project no phase progress and no favorite flag', () => {
    const created: IProjectListRow = upsertProjectRow([], form({ id: 0 })).rows[0];

    expect(created).toEqual({
      id: 1,
      name: 'Roof replacement',
      isFavorite: false,
      administratie: 'none',
      voorbereiding: 'none',
      ontwerp: 'none',
      uitvoering: 'none',
      revisie: 'none'
    });
  });

  it('trims surrounding whitespace from the project name', () => {
    const created: IProjectListRow = upsertProjectRow([], form({ id: 0, buildingName: '  Spaced  ' })).rows[0];

    expect(created.name).toBe('Spaced');
  });

  it('leaves the original list untouched', () => {
    const original: IProjectListRow[] = [row({ id: 1 })];

    upsertProjectRow(original, form({ id: 0 }));

    expect(original).toHaveLength(1);
  });
});

describe('upsertProjectRow - updating', () => {
  it('renames the matching row instead of appending', () => {
    const result = upsertProjectRow([row({ id: 1, name: 'Old name' })], form({ id: 1, buildingName: 'New name' }));

    expect(result.rows).toHaveLength(1);
    expect(result.rows[0].name).toBe('New name');
  });

  it('reports the id it updated', () => {
    expect(upsertProjectRow([row({ id: 9 })], form({ id: 9 })).id).toBe(9);
  });

  // The phase circles and the favorite star are listing state the form never edits, so an
  // update must not reset them.
  it('preserves phase progress and the favorite flag', () => {
    const existing: IProjectListRow = row({
      id: 1,
      isFavorite: true,
      administratie: 'success',
      ontwerp: 'danger',
      maintenance: 'success'
    });

    const updated: IProjectListRow = upsertProjectRow([existing], form({ id: 1 })).rows[0];

    expect(updated.isFavorite).toBe(true);
    expect(updated.administratie).toBe('success');
    expect(updated.ontwerp).toBe('danger');
    expect(updated.maintenance).toBe('success');
  });

  it('keeps the row in its original position', () => {
    const rows: IProjectListRow[] = [row({ id: 1 }), row({ id: 2 }), row({ id: 3 })];

    const result = upsertProjectRow(rows, form({ id: 2, buildingName: 'Middle' }));

    expect(result.rows[1].name).toBe('Middle');
  });

  it('adds it when the id is not in the list, rather than losing the project', () => {
    const result = upsertProjectRow([row({ id: 1 })], form({ id: 42, buildingName: 'Orphan' }));

    expect(result.rows).toHaveLength(2);
    expect(result.id).toBe(42);
    expect(result.rows[0].id).toBe(42);
  });
});
