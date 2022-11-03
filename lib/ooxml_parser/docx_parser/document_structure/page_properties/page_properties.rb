# frozen_string_literal: true

require_relative 'document_grid'
require_relative 'footnote_properties'
require_relative 'page_size'
require_relative 'page_margins'
require_relative 'columns'
require_relative 'note'
require_relative 'page_properties/header_footer_reference'
require_relative 'page_properties/page_borders'

module OoxmlParser
  # Class for data of PageProperties
  class PageProperties < OOXMLDocumentObject
    attr_accessor :type, :size, :margins, :document_grid, :num_type, :form_prot, :text_direction, :page_borders, :columns,
                  :notes
    # @return [FootnoteProperties] properties of footnote
    attr_accessor :footnote_properties
    # @return [True, False] Specifies whether the section should have a different header and footer for its first page
    attr_reader :title_page

    def initialize(parent: nil)
      @notes = []
      super
    end

    # Parse PageProperties data
    # @param [Nokogiri::XML:Element] node with PageProperties data
    # @return [PageProperties] value of PageProperties data
    def parse(node, default_paragraph, default_character)
      node.xpath('*').each do |pg_size_subnode|
        case pg_size_subnode.name
        when 'pgSz'
          @size = PageSize.new.parse(pg_size_subnode)
        when 'pgBorders'
          page_borders_object = PageBorders.new(parent: self).parse(pg_size_subnode)
          page_borders = Borders.new
          page_borders.display = page_borders_object.display if page_borders_object.display
          page_borders.offset_from = page_borders_object.offset_from if page_borders_object.offset_from
          pg_size_subnode.xpath('w:bottom').each do |bottom|
            page_borders.bottom = BordersProperties.new(parent: page_borders).parse(bottom)
          end
          pg_size_subnode.xpath('w:left').each do |left|
            page_borders.bottom = BordersProperties.new(parent: page_borders).parse(left)
          end
          pg_size_subnode.xpath('w:top').each do |top|
            page_borders.bottom = BordersProperties.new(parent: page_borders).parse(top)
          end
          pg_size_subnode.xpath('w:right').each do |right|
            page_borders.bottom = BordersProperties.new(parent: page_borders).parse(right)
          end
          @page_borders = page_borders
        when 'type'
          @type_object = ValuedChild.new(:string, parent: self).parse(pg_size_subnode)
          @type = @type_object.value
        when 'pgMar'
          @margins = PageMargins.new(parent: self).parse(pg_size_subnode)
        when 'pgNumType'
          @num_type = pg_size_subnode.attribute('fmt').value unless pg_size_subnode.attribute('fmt').nil?
        when 'formProt'
          form_prot_object = ValuedChild.new(:string, parent: self).parse(pg_size_subnode)
          @form_prot = form_prot_object.value
        when 'textDirection'
          text_directon_object = ValuedChild.new(:string, parent: self).parse(pg_size_subnode)
          @text_direction = text_directon_object.value
        when 'docGrid'
          @document_grid = DocumentGrid.new(parent: self).parse(pg_size_subnode)
        when 'titlePg'
          @title_page = option_enabled?(pg_size_subnode)
        when 'cols'
          @columns = Columns.new.parse(pg_size_subnode)
        when 'headerReference', 'footerReference'
          reference = HeaderFooterReference.new(parent: self).parse(pg_size_subnode)
          target = root_object.get_link_from_rels(reference.id)
          root_object.add_to_xmls_stack("word/#{target}")
          note = Note.new.parse(default_paragraph: default_paragraph,
                                default_character: default_character,
                                target: target,
                                assigned_to: reference.type,
                                type: File.basename(target).sub('.xml', ''),
                                parent: self)
          @notes << note
          root_object.xmls_stack.pop
        when 'footnotePr'
          @footnote_properties = FootnoteProperties.new(parent: self).parse(pg_size_subnode)
        end
      end
      self
    end
  end
end
