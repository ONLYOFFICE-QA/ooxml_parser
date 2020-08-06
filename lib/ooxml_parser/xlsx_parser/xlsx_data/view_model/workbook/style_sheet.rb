# frozen_string_literal: true

require_relative 'style_sheet/cell_xfs'
require_relative 'style_sheet/fills'
require_relative 'style_sheet/fonts'
require_relative 'style_sheet/number_formats'
require_relative 'style_sheet/xlsx_borders'
module OoxmlParser
  # Parsing file styles.xml
  class StyleSheet < OOXMLDocumentObject
    # @return [NumberFormats] number formats
    attr_accessor :number_formats
    # @return [Fonts] fonts
    attr_accessor :fonts
    # @return [Fills] fills
    attr_accessor :fills
    # @return [CellXfs] Cell XFs
    attr_reader :cell_xfs
    # @return [XlsxBorders] Cell XFs
    attr_reader :borders

    def initialize(parent: nil)
      @number_formats = NumberFormats.new(parent: self)
      @fonts = Fonts.new(parent: self)
      @fills = Fills.new(parent: self)
      super
    end

    # Parse StyleSheet object
    # @return [StyleSheet] result of parsing
    def parse
      doc = parse_xml("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/styles.xml")
      doc.root.xpath('*').each do |node_child|
        case node_child.name
        when 'numFmts'
          @number_formats.parse(node_child)
        when 'fonts'
          @fonts.parse(node_child)
        when 'fills'
          @fills.parse(node_child)
        when 'cellXfs'
          @cell_xfs = CellXfs.new(parent: self).parse(node_child)
        when 'borders'
          @borders = XlsxBorders.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
