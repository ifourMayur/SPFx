import { IProjectService } from '../../../../services/ProjectService';

export interface IProjectProps {
  /**
   * Injected by the web part through `ServiceFactory`, ready for this component tree to
   * call. Keeps API access out of the component itself.
   */
  projectService: IProjectService;
}
