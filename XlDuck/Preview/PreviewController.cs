// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck.Preview;

/// <summary>
/// Controls preview loading with debounce and serial queue.
/// </summary>
public class PreviewController : IDisposable
{
    private const int DebounceMs = 200;

    private readonly PreviewPane _pane;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private CancellationTokenSource? _cts;
    private bool _disposed;

    public PreviewController(PreviewPane pane)
    {
        _pane = pane;
    }

    /// <summary>
    /// Called when cell selection changes. Debounces and loads preview.
    /// </summary>
    public void OnSelectionChanged(string? cellValue)
    {
        if (_disposed) return;

        // Cancel any pending work
        _cts?.Cancel();
        _cts?.Dispose();
        _cts = new CancellationTokenSource();

        var token = _cts.Token;

        // Start debounced load
        _ = LoadPreviewDebouncedAsync(cellValue, token);
    }

    private async Task LoadPreviewDebouncedAsync(string? cellValue, CancellationToken token)
    {
        try
        {
            // Debounce
            await Task.Delay(DebounceMs, token);

            // Wait for gate (serial queue)
            await _gate.WaitAsync(token);

            try
            {
                // Check if still valid
                if (token.IsCancellationRequested) return;

                // Load preview model (runs on thread pool to avoid blocking UI)
                var model = await Task.Run(() => PreviewDataProvider.GetPreview(cellValue), token);

                // Check again after work completed
                if (token.IsCancellationRequested) return;

                // Update pane on UI thread
                if (_pane.InvokeRequired)
                {
                    _pane.Invoke(() => _pane.SetState(model));
                }
                else
                {
                    _pane.SetState(model);
                }
            }
            finally
            {
                _gate.Release();
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when selection changes rapidly
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewController] Error: {ex.Message}");

            // Show error in pane
            try
            {
                var errorModel = new ErrorPreviewModel
                {
                    Title = "Preview Error",
                    Message = ex.Message
                };

                if (_pane.InvokeRequired)
                {
                    _pane.Invoke(() => _pane.SetState(errorModel));
                }
                else
                {
                    _pane.SetState(errorModel);
                }
            }
            catch
            {
                // Ignore errors showing error
            }
        }
    }

    /// <summary>
    /// Force immediate refresh (no debounce).
    /// </summary>
    public void RefreshNow(string? cellValue)
    {
        if (_disposed) return;

        // Cancel any pending work
        _cts?.Cancel();
        _cts?.Dispose();
        _cts = new CancellationTokenSource();

        var token = _cts.Token;

        // Load immediately (still uses serial queue)
        _ = LoadPreviewImmediateAsync(cellValue, token);
    }

    private async Task LoadPreviewImmediateAsync(string? cellValue, CancellationToken token)
    {
        try
        {
            await _gate.WaitAsync(token);

            try
            {
                if (token.IsCancellationRequested) return;

                var model = await Task.Run(() => PreviewDataProvider.GetPreview(cellValue), token);

                if (token.IsCancellationRequested) return;

                if (_pane.InvokeRequired)
                {
                    _pane.Invoke(() => _pane.SetState(model));
                }
                else
                {
                    _pane.SetState(model);
                }
            }
            finally
            {
                _gate.Release();
            }
        }
        catch (OperationCanceledException)
        {
            // Expected
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewController] RefreshNow error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _cts?.Cancel();
        _cts?.Dispose();
        _gate.Dispose();
    }
}
