# Simple PowerShell HTTP Server for PM Dashboard
$port = 3000
$localDir = "C:\Users\it-support706.dsi\.gemini\antigravity\scratch\pm-dashboard"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    $listener.Start()
    Write-Output "PowerShell Web Server started successfully!"
    Write-Output "Open your browser and navigate to: http://localhost:$port/"
    Write-Output "Press Ctrl+C or kill this background task to stop the server."
} catch {
    Write-Error "Failed to start HTTP listener: $_"
    exit 1
}

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        $urlPath = $request.Url.LocalPath
        # Default document
        if ($urlPath -eq "/") {
            $urlPath = "/index.html"
        }
        
        # Build safe local file path
        # Remove leading slash and combine
        $relativeFile = $urlPath.TrimStart('/')
        $filePath = Join-Path $localDir $relativeFile
        
        if (Test-Path $filePath -PathType Leaf) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            
            # Determine MIME type based on extension
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            $mime = "application/octet-stream"
            
            if ($ext -eq ".html" -or $ext -eq ".htm") { $mime = "text/html; charset=utf-8" }
            elseif ($ext -eq ".css") { $mime = "text/css; charset=utf-8" }
            elseif ($ext -eq ".js") { $mime = "text/javascript; charset=utf-8" }
            elseif ($ext -eq ".json") { $mime = "application/json; charset=utf-8" }
            elseif ($ext -eq ".png") { $mime = "image/png" }
            elseif ($ext -eq ".jpg" -or $ext -eq ".jpeg") { $mime = "image/jpeg" }
            elseif ($ext -eq ".svg") { $mime = "image/svg+xml; charset=utf-8" }
            
            $response.ContentType = $mime
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            # File not found
            $response.StatusCode = 404
            $errContent = [System.Text.Encoding]::UTF8.GetBytes("<h1>404 File Not Found</h1><p>Path: $urlPath</p>")
            $response.ContentType = "text/html; charset=utf-8"
            $response.ContentLength64 = $errContent.Length
            $response.OutputStream.Write($errContent, 0, $errContent.Length)
        }
    } catch {
        # Catch individual request handling errors to prevent the server from crashing
        Write-Warning "Error handling request: $_"
    } finally {
        if ($response) {
            try {
                $response.OutputStream.Close()
            } catch {}
        }
    }
}
