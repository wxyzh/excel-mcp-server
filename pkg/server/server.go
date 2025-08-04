package server

import (
	"runtime"

	"github.com/mark3labs/mcp-go/server"
	"github.com/wxyzh/excel-mcp-server/pkg/tools"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New(version string) *ExcelServer {
	s := &ExcelServer{}
	s.server = server.NewMCPServer(
		"excel-mcp-server",
		version,
	)
	tools.AddExcelDescribeSheetsTool(s.server)
	tools.AddExcelReadSheetTool(s.server)
	if runtime.GOOS == "windows" {
		tools.AddExcelScreenCaptureTool(s.server)
	}
	tools.AddExcelWriteToSheetTool(s.server)
	tools.AddExcelCreateTableTool(s.server)
	tools.AddExcelCopySheetTool(s.server)
	tools.AddExcelFormatRangeTool(s.server)
	return s
}

func (s *ExcelServer) Start() error {
	return server.ServeStdio(s.server)
}

func (s *ExcelServer) GetMCPServer() *server.MCPServer {
    return s.server
}