# frozen_string_literal: true

require_relative 'presentation/comment_authors'
require_relative 'presentation/presentation_comments'
require_relative 'presentation/presentation_helpers'
require_relative 'presentation/presentation_theme'
require_relative 'presentation/slide'
require_relative 'presentation/slide_master_file'
require_relative 'presentation/slide_masters_helper'
require_relative 'presentation/slide_layout_file'
require_relative 'presentation/slide_layouts_helper'
require_relative 'presentation/slide_size'
require_relative 'presentation/table_styles'
module OoxmlParser
  # Basic class for all parsed pptx data
  class Presentation < CommonDocumentStructure
    include PresentationHelpers
    include SlideLayoutsHelper
    include SlideMastersHelper
    attr_accessor :slides, :theme, :slide_size
    # @return [Relationships] relationships of presentation
    attr_accessor :relationships
    # @return [TableStyles] table styles data
    attr_accessor :table_styles
    # @return [CommentAuthors] authors of presentation
    attr_reader :comment_authors
    # @return [PresentationComments] comments of presentation
    attr_reader :comments
    # @return [Array<SlideMasterFile>] list of slide master
    attr_reader :slide_masters
    # @return [Array<SlideLayout>] list of slide layouts
    attr_reader :slide_layouts

    def initialize(params = {})
      @slides = []
      @comments = []
      @slide_masters = []
      @slide_layouts = []
      super
    end

    # Parse data of presentation
    # @return [Presentation] parsed presentation
    def parse
      @content_types = ContentTypes.new(parent: self).parse
      OOXMLDocumentObject.root_subfolder = 'ppt/'
      OOXMLDocumentObject.xmls_stack = []
      OOXMLDocumentObject.add_to_xmls_stack('ppt/presentation.xml')
      doc = parse_xml(OOXMLDocumentObject.current_xml)
      @theme = PresentationTheme.parse('ppt/theme/theme1.xml')
      @table_styles = TableStyles.new(parent: self).parse
      @comment_authors = CommentAuthors.new(parent: self).parse
      @comments = PresentationComments.new(parent: self).parse
      presentation_node = doc.search('//p:presentation').first
      presentation_node.xpath('*').each do |presentation_node_child|
        case presentation_node_child.name
        when 'sldSz'
          @slide_size = SlideSize.new(parent: self).parse(presentation_node_child)
        when 'sldIdLst'
          presentation_node_child.xpath('p:sldId').each do |silde_id_node|
            slide_id = silde_id_node.attr('r:id')
            @slides << Slide.new(parent: self,
                                 xml_path: "#{OOXMLDocumentObject.root_subfolder}/#{OOXMLDocumentObject.get_link_from_rels(slide_id)}")
                            .parse
          end
        end
      end
      OOXMLDocumentObject.xmls_stack.pop
      @relationships = Relationships.new(parent: self).parse_file("#{OOXMLDocumentObject.path_to_folder}/ppt/_rels/presentation.xml.rels")
      parse_slide_layouts
      parse_slide_masters
      self
    end
  end
end
