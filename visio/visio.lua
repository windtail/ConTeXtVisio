
module("visio", package.seeall)

require "luacom"
local iconv = require "iconv"
require "lfs"
require "md5"

local standaloneDebug = (context == nil)

if standaloneDebug then
	context = function(...)
		print(string.format(...))
	end
end

function writeStatus(msg)
	msg = msg:gsub("\\","/") -- 确保消息中没有不合法的TeX命令!
	msg = msg:gsub("#", "")

	msg = msg:gsub("[%%\\#{}]", function(pattern)
		local subst = {["%"]="@",["\\"]="/",["#"]="$", ["{"]="!", ["}"]="!"}
		return subst[pattern]
	end)

	context("\\writestatus{visio status}{" .. msg .. "}")
end

--~ -------------------------------------------
--~ visio constants
local visOpenRW = 0x20
local visTypeShape = 3
local visBBoxUprightWH = 0x1
local visSectionObject = 1
local visRowPage = 10
local visPageWidth = 0
local visPageHeight = 1
local visPageDrawSizeType = 6
local visRowPrintProperties = 25
local visPrintPropertiesLeftMargin = 0
local visPrintPropertiesRightMargin = 1
local visPrintPropertiesTopMargin = 2
local visPrintPropertiesBottomMargin = 3
local visPrintPropertiesPaperSource = 18
local visPrintPropertiesCenterX = 8
local visPrintPropertiesCenterY = 9
local visFixedFormatPDF = 1
local visDocExIntentScreen = 0
local visPrintCurrentPage = 2

--~ -------------------------------------------
--~ 编码转换
--~ 编写TeX源文件时一般用UTF-8格式
--~ 系统用什么编码由用户设置，默认GBK

local function createIconv(to, from)
	local cd = iconv.new(to, from)
	return function(txt)
		return cd:iconv(txt)
	end
end

local sourceEncoding = "utf-8"
local systemEncoding = "gbk"
local L = createIconv(sourceEncoding, systemEncoding)
local Z = createIconv(systemEncoding, sourceEncoding)

function setupEncoding(sys, src)
	if sys ~= nil and #sys > 0 then
		systemEncoding = sys
	end

	if src ~= nil and #src > 0 then
		sourceEncoding = src
	end

	if sourceEncoding == systemEncoding then
		local echo = function(txt) return txt end
		L = echo
		Z = echo
	else
		L = createIconv(sourceEncoding, systemEncoding)
		Z = createIconv(systemEncoding, sourceEncoding)
	end
end

--~ -------------------------------------------
--~ File system related functions

local function isFileExists(path)
	return lfs.attributes(path, "dev") ~= nil
end

local function isAbsolutePath(filename)
	if filename:find("%a:\\") == 1 then
		return true
	else
		return false
	end
end

local function isVisioUpdated(visioPath, pdfPath)
	local visioTime = assert(lfs.attributes(visioPath, "modification"))

	if not isFileExists(pdfPath) then return true end

	local pdfTime = assert(lfs.attributes(pdfPath, "modification"))

	return visioTime > pdfTime
end

local function createDirectoryRecursively(dir)
	local i,j
	local status, err

	assert(isAbsolutePath(dir))

	j = 3 -- ignore driver letters
	while true do
		i,j = dir:find("[^\\]+", j+1)
		if i then
			local subdir = dir:sub(1,j)
			if not isFileExists(subdir) then
				writeStatus("Creating directory " .. subdir)
				status, err = lfs.mkdir(subdir)
				if not status then
					return false, err
				end
			end
		else
			return true
		end
	end
end

--~ -------------------------------------------
--~ visio convertion

--~ may throw errors
--~ return width, height in mm
local function vis2pdf(visioApp, visioPath, pdfPath, margin)
	local l, b, r, t, width, height

	local visioDoc = visioApp.Documents:OpenEx(visioPath, visOpenRW)
	local visioWnd = visioApp.ActiveWindow
	local visioPage = visioApp.ActivePage
	local visioSheet = visioApp.ActivePage.PageSheet
	local visioCell

	visioWnd:SelectAll()
	l, b, r, t = visioWnd.Selection:BoundingBox(visTypeShape+visBBoxUprightWH)

	width = (r - l) * 25.4 + margin.left + margin.right -- 1inch = 25.4mm
	height = (t - b) * 25.4 + margin.top + margin.bottom

	visioPage.Background = false
	visioPage.BackPage = ""

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPage, visPageWidth)
	visioCell.FormulaU = width .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPage, visPageHeight)
	visioCell.FormulaU = height .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPage, visPageDrawSizeType)
	visioCell.FormulaU = "1"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPage, 38)
	visioCell.FormulaU = "2"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesLeftMargin)
	visioCell.FormulaU = margin.left .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesRightMargin)
	visioCell.FormulaU = margin.right .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesTopMargin)
	visioCell.FormulaU = margin.top .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesBottomMargin)
	visioCell.FormulaU = margin.bottom .. " mm"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPaperSource)
	visioCell.FormulaU = "15"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesCenterX)
	visioCell.FormulaU = "1"

	visioCell = visioSheet:getCellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesCenterY)
	visioCell.FormulaU = "1"

	visioWnd:SelectAll()
	l, b, r, t = visioApp.ActiveWindow.Selection:BoundingBox(visTypeShape+visBBoxUprightWH)
	visioWnd.Selection:Move(-l * 25.4 + margin.left, -b * 25.4 + margin.bottom, "mm")

	visioDoc:ExportAsFixedFormat(visFixedFormatPDF, pdfPath, visDocExIntentScreen, visPrintCurrentPage,
		1, -1, false, true, false, false, true)

	-- 关闭但不保存文档
	visioDoc.Saved = true
	visioDoc:Close()

	visioDoc = nil
	visioWnd = nil
	visioPage = nil
	visioSheet = nil
	visioCell = nil

	return width, height
end

-- catch errors
-- 成功返回 true, width, height
-- 失败返回 false, errMessage
local function safeVis2pdf(visioPath, pdfPath, margin)
	local width, height

	local status, err = pcall(function ()
		local visioApp = luacom.CreateObject("Visio.InvisibleApp")
		-- visioApp.Visible = true -- for debug

		-- 保证即使错误发生也要关闭visio
		local status, err = pcall(function()
			width, height = vis2pdf(visioApp, visioPath, pdfPath, margin)
		end)

		-- 下面这句话基本上我认为不会发生错误 {
		for i=1,visioApp.Documents.Count do
			visioApp.Documents:Item(i).Saved = true
		end
		-- }

		visioApp:Quit()
		visioApp = nil

		if not status then
			error(err)
		end
	end)

	if not status then
		return false, L(err) -- convert err encoding
	else
		return true, width, height
	end
end

-- update pdf only when pdf file is older than visio file
-- visioPath and pdfPath MUST be converted to system encoding before calling this function
local function updatePdf(visioPath, pdfPath, margin)
	assert(visioPath)
	assert(pdfPath)
	assert(margin)

	-- 比较时间看是否需要转换
	if not isVisioUpdated(visioPath, pdfPath) then
		return
	end

	-- NOTE:
	-- the dimension returned are not used yet
	-- cause i think ConTeXt can determine the dimension of a PDF
	-- if it cannot, we have to introduce a temp file to store every PDF's
	-- dimension, when PDF need't update we have to find the dimension in
	-- the temp file
	local status, errOrWidth, height = safeVis2pdf(visioPath, pdfPath, margin)
	if not status then
		error(errOrWidth)
	end

	-- 等待转换完毕
	--local fso = luacom.CreateObject("Scripting.FileSystemObject")
	--while not fso:FileExists(pdfPath) do
	--end

	return errOrWidth, height
end

--~ -------------------------------------------
-- PDF output directory

-- the converted PDF files are put into the file
local pdfDir = luacom.GetCurrentDirectory()

-- map visio filename supplied by user to pdf filename
local pdfMap = {}

local function addPdf(visioFileName)
	local prefix = "p"
	pdfMap[visioFileName] = prefix .. md5.sumhexa(visioFileName) .. ".pdf"
	return pdfMap[visioFileName]
end

local function getPdfName(visioFileName)
	return pdfMap[visioFileName] or pdfMap[visioFileName .. ".vsd"] or "undefined"
end

-- use / instead of \\
local function setupPdfDir(dir)
	assert(dir and #dir ~= 0)

	-- figures.paths store the current paths!!
	-- a little slow, but this is not called frequently (usually only once)
	local paths = {}
	local excludedPaths = {}
	for _,p in ipairs(figures.localpaths) do
		excludedPaths[p] = true
	end
	excludedPaths[pdfDir] = true

	for _,p in ipairs(figures.paths) do
		if not excludedPaths[p] then
			table.insert(paths, p)
		end
	end

	dir = dir:gsub("/", "\\")

	if isAbsolutePath(dir) then
		pdfDir = dir
	else
		pdfDir = luacom.GetCurrentDirectory() .. "\\" .. dir
	end

	writeStatus("PDF output directory is set to " .. pdfDir)

	table.insert(paths, (pdfDir:gsub("\\", "/"))) -- suppress gsub return value to only one
	context("\\setupexternalfigures[directory={" .. table.concat(paths, ",") .."}]")
end

-- PDF path with windows' seperator "\\"
local function getPdfPath(visioFileName)
	return pdfDir .. "\\" .. addPdf(visioFileName)
end

--~ -------------------------------------------
-- VISIO search directories

local visioDirs = {}
local localVisioDirs = {".", "..", "image", "figure", "visio"}

-- VISIO path with windows' seperator "\\" and local system encoding
local function searchVisioPath(visioFileName)
	if visioFileName:sub(-4) ~= ".vsd" then
		visioFileName = visioFileName .. ".vsd"
	end

	visioFileName = Z(visioFileName:gsub("/", "\\")) -- to system encoding

	if isAbsolutePath(visioFileName) then
		return visioFileName
	end

	-- search local directory
	local curDir = luacom.GetCurrentDirectory()
	local path

	for _,subdir in ipairs(localVisioDirs) do
		path = curDir .. "\\" .. subdir .. "\\" .. visioFileName
		if isFileExists(path) then
			return path
		end
	end

	-- search global directory
	for _,dir in ipairs(visioDirs) do
		path = dir .. "\\" .. visioFileName
		if isFileExists(path) then
			return path
		end
	end

	if standaloneDebug then
		assert(false, visioFileName .. " not found!")
	end
	return nil
end

-- @param dirs for example c:/aa,d:/bbb
local function setupVisioDirs(dirs)
	assert(dirs and #dirs ~= 0)

	visioDirs = {} -- clear old dirs

	dirs = dirs:gsub("/", "\\")
	for dir in dirs:gmatch("[^,]+") do
		table.insert(visioDirs, dir)
	end
end

--~ -------------------------------------------

function useVisio(visioFileName, leftMargin, rightMargin, topMargin, bottomMargin)
	leftMargin = leftMargin or 0
	rightMargin = rightMargin or 0
	topMargin = topMargin or 0
	bottomMargin = bottomMargin or 0

	margin = {left=leftMargin, right=rightMargin, top=topMargin, bottom=bottomMargin}
	writeStatus("left=" .. margin.left .. " right=" .. margin.right .. " top=" .. margin.top .. " bottom=" .. margin.bottom)

	local visioPath = assert(searchVisioPath(visioFileName), visioFileName .. " not found!")
	local pdfPath = getPdfPath(visioFileName)

	-- delay the creatation of PDF output directory until necessary
	if not isFileExists(pdfDir) then
		assert(createDirectoryRecursively(pdfDir))
	end

	local w, h = updatePdf(visioPath, pdfPath, margin)
	context("\\useexternalfigure[%s]", getPdfName(visioFileName))

	--[[ not yet implemented
	if w and h then
		context("[%.4fmm][%.4fmm]", w, h)
	end--]]
end

function visio(visioFileName)
	context("\\externalfigure[%s]", getPdfName(visioFileName))
end

function setupDirectory(category, list)

	if list == nil or #list == 0 then return end

	if category == "pdf" then
		setupPdfDir(list)
	elseif category == "visio" then
		setupVisioDirs(list)
	else
		writeStatus("unknown #1 for setupvisiodirectory, ignored")
	end
end

if standaloneDebug then
	setupDirectory("visio", "D:/ConTeXt-note/head")
	useVisio("中文文件名啊")
	visio("中文文件名啊")
end

