// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;

namespace XlDuck.Core;

public sealed record ComputationStatus(string State, string Handle, string Kind, DateTimeOffset CreatedAt,
    DateTimeOffset? StartedAt, DateTimeOffset? CompletedAt, long ElapsedMilliseconds,
    ImmutableArray<string> Dependencies, long? Rows, ComputationFailure? Failure);
