import { GS, Student, fetchData } from '@lib';
import { NamedRange, sheets } from '@utils/constants';

const getMembers = (present?: boolean): Student[] =>
  fetchData(GS.ss.getRangeByName(NamedRange.ProjectMembers), {
    filter: row => (present === undefined ? true : present === !!row[4]),
    map: row =>
      new Student({
        name: row[1],
        nUsp: '',
        nickname: row[2],
        email: row[3],
      }),
  });

export const getAllMembers = (): Student[] => getMembers();

export const getPresentMembers = (): Student[] => getMembers(true);

export const getMissingMembers = (): Student[] => getMembers(false);

export const getBoardOfDirectors = (): Student[] =>
  fetchData(sheets.boardOfDirectors, {
    map: row =>
      new Student({
        name: row[0],
        nUsp: '',
        nickname: row[1],
        email: row[2],
      }),
  });
