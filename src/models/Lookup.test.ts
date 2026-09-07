import { toLookupOptions } from './Lookup';

describe('toLookupOptions', () => {
  it('maps the API LookupItem shape onto dropdown options', () => {
    const result = toLookupOptions([
      { id: 3, name: 'Renovation' },
      { id: 4, name: 'Maintenance only' }
    ]);

    expect(result).toEqual([
      { id: '3', name: 'Renovation' },
      { id: '4', name: 'Maintenance only' }
    ]);
  });

  // The reference binds these with `new SelectList(list, "Id", "Name")`, so the numeric Id
  // is the option value; it becomes a string because that is what a <select> value is.
  it('stringifies the numeric id', () => {
    expect(toLookupOptions([{ id: 42, name: 'x' }])[0].id).toBe('42');
  });

  it('prefers strId when the item carries one', () => {
    expect(toLookupOptions([{ id: 0, strId: 'abc', name: 'x' }])[0].id).toBe('abc');
  });

  it('falls back to the numeric id when strId is blank', () => {
    expect(toLookupOptions([{ id: 7, strId: '   ', name: 'x' }])[0].id).toBe('7');
  });

  it('returns an empty list for undefined, so a failed load renders an empty dropdown', () => {
    expect(toLookupOptions(undefined)).toEqual([]);
  });

  it('returns an empty list for null', () => {
    expect(toLookupOptions(null)).toEqual([]);
  });

  it('returns an empty list for a payload that is not an array', () => {
    expect(toLookupOptions({ unexpected: true })).toEqual([]);
  });

  // System.Text.Json writes `null` rather than omitting keys, so a null name must not
  // become the string "null" in a dropdown.
  it('renders a null name as an empty label rather than "null"', () => {
    expect(toLookupOptions([{ id: 1, name: null }])[0].name).toBe('');
  });

  it('skips entries that carry no usable identifier', () => {
    expect(toLookupOptions([{ id: 1, name: 'keep' }, undefined, null, { name: 'no id' }])).toEqual([
      { id: '1', name: 'keep' }
    ]);
  });
});
