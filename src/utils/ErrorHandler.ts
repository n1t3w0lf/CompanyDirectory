import { IErrorDetails } from '../models/IUserProfile';
import { Constants } from '../models/Constants';

/**
 * Centralized error handling utility
 */
export class ErrorHandler {
  /**
   * Parse and categorize errors
   */
  public static parseError(error: unknown, context?: string): IErrorDetails {
    const timestamp = new Date();

    // Handle Graph API errors
    if (this.isGraphError(error)) {
      const graphError = error as { statusCode?: number | string; code?: number | string };
      const statusCode = graphError.statusCode || graphError.code;

      switch (statusCode) {
        case 401:
        case 403:
          return {
            message: Constants.ERROR_PERMISSION,
            code: statusCode.toString(),
            timestamp,
            context
          };
        case 429:
          return {
            message: 'Too many requests. Please wait a moment and try again.',
            code: '429',
            timestamp,
            context
          };
        case 404:
          return {
            message: 'Resource not found.',
            code: '404',
            timestamp,
            context
          };
        case 500:
        case 503:
          return {
            message: 'Service temporarily unavailable. Please try again later.',
            code: statusCode.toString(),
            timestamp,
            context
          };
        default:
          return {
            message: Constants.ERROR_GRAPH_API,
            code: statusCode?.toString() || 'GRAPH_ERROR',
            timestamp,
            context
          };
      }
    }

    // Handle network errors
    if (error instanceof TypeError && error.message.includes('fetch')) {
      return {
        message: Constants.ERROR_NETWORK,
        code: 'NETWORK_ERROR',
        timestamp,
        context
      };
    }

    // Handle standard errors
    if (error instanceof Error) {
      return {
        message: error.message || Constants.ERROR_GENERIC,
        code: 'ERROR',
        timestamp,
        context
      };
    }

    // Unknown error type
    return {
      message: Constants.ERROR_GENERIC,
      code: 'UNKNOWN_ERROR',
      timestamp,
      context
    };
  }

  /**
   * Check if error is from Microsoft Graph API
   */
  private static isGraphError(error: unknown): boolean {
    return (
      typeof error === 'object' &&
      error !== null &&
      ('statusCode' in error || 'code' in error)
    );
  }

  /**
   * Log error to console (can be extended to log to Application Insights)
   */
  public static logError(error: IErrorDetails): void {
    console.error('[People Directory Error]', {
      message: error.message,
      code: error.code,
      timestamp: error.timestamp,
      context: error.context
    });

    // TODO: Integrate with Application Insights or other logging service
    // Example:
    // appInsights.trackException({ exception: new Error(error.message), properties: error });
  }

  /**
   * Get user-friendly error message
   */
  public static getUserMessage(error: unknown, context?: string): string {
    const errorDetails = this.parseError(error, context);
    this.logError(errorDetails);
    return errorDetails.message;
  }
}
