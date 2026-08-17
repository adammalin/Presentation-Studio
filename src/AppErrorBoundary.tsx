import { Component, type ErrorInfo, type ReactNode } from "react";

interface AppErrorBoundaryState {
  failed: boolean;
}

export default class AppErrorBoundary extends Component<{ children: ReactNode }, AppErrorBoundaryState> {
  state: AppErrorBoundaryState = { failed: false };

  static getDerivedStateFromError(): AppErrorBoundaryState {
    return { failed: true };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("Presentation Studio renderer error", error, info.componentStack);
  }

  render() {
    if (!this.state.failed) return this.props.children;
    return <main className="app-error-state" role="alert">
      <span>Presentation Studio</span>
      <h1>This workspace could not be displayed.</h1>
      <p>Your project and imported source files were not changed. Reload the local app to restore the workspace.</p>
      <button type="button" onClick={() => window.location.reload()}>Reload Presentation Studio</button>
    </main>;
  }
}
