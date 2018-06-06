require_relative 'presentation/presentation_comment'
require_relative 'presentation/presentation_helpers'
require_relative 'presentation/presentation_theme'
require_relative 'presentation/slide'
require_relative 'presentation/slide_size'
require_relative 'presentation/table_styles'
module OoxmlParser
  class Presentation < CommonDocumentStructure
    include PresentationHelpers
    attr_accessor :slides, :theme, :slide_size, :comments
    # @return [Relationships] relationships of presentation
    attr_accessor :relationships
    # @return [TableStyles] table styles data
    attr_accessor :table_styles

    def initialize(params = {})
      @slides = []
      @comments = []
      super
    end

    def parse
      OOXMLDocumentObject.root_subfolder = 'ppt/'
      OOXMLDocumentObject.xmls_stack = []
      OOXMLDocumentObject.add_to_xmls_stack('ppt/presentation.xml')
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.current_xml))
      @theme = PresentationTheme.parse('ppt/theme/theme1.xml')
      @table_styles = TableStyles.new(parent: self).parse
      presentation_node = doc.search('//p:presentation').first
      presentation_node.xpath('*').each do |presentation_node_child|
        case presentation_node_child.name
        when 'sldSz'
          @slide_size = SlideSize.new(parent: self).parse(presentation_node_child)
        when 'sldIdLst'
          presentation_node_child.xpath('p:sldId').each do |silde_id_node|
            id = nil
            silde_id_node.attribute_nodes.select { |node| id = node.to_s if node.namespace && node.namespace.prefix == 'r' }
            @slides << Slide.new(parent: self,
                                 xml_path: "#{OOXMLDocumentObject.root_subfolder}/#{OOXMLDocumentObject.get_link_from_rels(id)}")
                            .parse
          end
        end
      end
      @comments = PresentationComment.parse_list
      OOXMLDocumentObject.xmls_stack.pop
      @relationships = Relationships.parse_rels("#{OOXMLDocumentObject.path_to_folder}/ppt/_rels/presentation.xml.rels")
      self
    end
  end
end
