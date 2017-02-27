require_relative 'style_sheet/fills'
require_relative 'style_sheet/fonts'
require_relative 'style_sheet/number_formats'
module OoxmlParser
  # Parsing file styles.xml
  class StyleSheet < OOXMLDocumentObject
    # @return [NumberFormats] number formats
    attr_accessor :number_formats
    # @return [Fonts] fonts
    attr_accessor :fonts
    # @return [Fills] fills
    attr_accessor :fills

    def initialize(parent: nil)
      @number_formats = NumberFormats.new(parent: self)
      @fonts = Fonts.new(parent: self)
      @fills = Fills.new(parent: self)
      @parent = parent
    end

    def parse
      doc = Nokogiri::XML(File.open("#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/styles.xml"))
      doc.root.xpath('*').each do |node_child|
        case node_child.name
        when 'numFmts'
          @number_formats.parse(node_child)
        when 'fonts'
          @fonts.parse(node_child)
        when 'fills'
          @fills.parse(node_child)
        end
      end
      self
    end
  end
end
