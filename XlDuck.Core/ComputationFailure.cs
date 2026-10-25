// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Core;

/// <summary>Stable, public failure categories. Messages contain no SQL excerpts or credentials.</summary>
public sealed record ComputationFailure(string Code, string Message, string? Advice = null)
{
    public string Display => $"[{Code}] {Message}" + (Advice is null ? "" : " " + Advice);
    public static ComputationFailure From(Exception error)
    {
        while (error is AggregateException aggregate && aggregate.InnerExceptions.Count == 1) error = aggregate.InnerExceptions[0];
        return error switch
        {
            ComputationException computation => computation.Failure,
            OperationCanceledException => new("cancelled", "Computation cancelled.", "Use Retry failed results to try again."),
            ArgumentException => new("invalid_input", error.Message),
            IOException => new("worker_unavailable", "The XlDuck worker is unavailable.", "Restart XlDuck for this workbook."),
            TimeoutException => new("timeout", "The operation exceeded its time limit.", "Reduce the query or retry when the workbook is idle."),
            _ => new("execution", error.Message)
        };
    }
}

public sealed class ComputationException(ComputationFailure failure) : InvalidOperationException(failure.Display)
{
    public ComputationFailure Failure { get; } = failure;
}
