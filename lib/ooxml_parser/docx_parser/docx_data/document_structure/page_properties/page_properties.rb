require_relative 'document_grid'
require_relative 'size'
require_relative 'page_margins'
require_relative 'columns'
require_relative 'note'
module OoxmlParser
  class PageProperties < OOXMLDocumentObject
    attr_accessor :type, :size, :margins, :document_grid, :num_type, :form_prot, :text_direction, :page_borders, :columns,
                  :notes

    def initialize
      @type = nil
      @size = nil
      @margins = nil
      @document_grid = nil
      @num_type = nil
      @form_prot = nil
      @text_direction = nil
      @columns = nil
      @notes = []
    end

    def self.parse(sect_pr, default_paragraph, default_character)
      page_properties = PageProperties.new
      sect_pr.xpath('w:pgSz').each do |pg_sz|
        page_properties.size = Size.parse(pg_sz)
      end
      sect_pr.xpath('w:pgBorders').each do |pg_borders|
        page_borders = Borders.new
        unless pg_borders.attribute('display').nil?
          page_borders.display = pg_borders.attribute('display').value.to_sym
        end
        unless pg_borders.attribute('offsetFrom').nil?
          page_borders.offset_from = pg_borders.attribute('offsetFrom').value.to_sym
        end
        pg_borders.xpath('w:bottom').each do |bottom|
          page_borders.bottom = BordersProperties.parse(bottom)
        end
        pg_borders.xpath('w:left').each do |left|
          page_borders.bottom = BordersProperties.parse(left)
        end
        pg_borders.xpath('w:top').each do |top|
          page_borders.bottom = BordersProperties.parse(top)
        end
        pg_borders.xpath('w:right').each do |right|
          page_borders.bottom = BordersProperties.parse(right)
        end
        page_properties.page_borders = page_borders
      end
      sect_pr.xpath('w:type').each do |type|
        page_properties.type = type.attribute('val').value
      end
      sect_pr.xpath('w:pgMar').each do |pg_mar|
        page_properties.margins = PageMargins.parse(pg_mar)
      end
      sect_pr.xpath('w:pgNumType').each do |pg_num_type|
        page_properties.num_type = pg_num_type.attribute('fmt').value unless pg_num_type.attribute('fmt').nil?
      end
      sect_pr.xpath('w:formProt').each do |form_prot|
        page_properties.form_prot = form_prot.attribute('val').value
      end
      sect_pr.xpath('w:textDirection').each do |text_direction|
        page_properties.text_direction = text_direction.attribute('val').value
      end
      sect_pr.xpath('w:docGrid').each do |doc_grid|
        page_properties.document_grid = DocumentGrid.parse(doc_grid)
      end
      sect_pr.xpath('w:cols').each do |columns_grid|
        page_properties.columns = Columns.parse(columns_grid)
      end
      sect_pr.xpath('w:headerReference').each do |header_reference|
        target = OOXMLDocumentObject.get_link_from_rels(header_reference.attribute('id').value)
        OOXMLDocumentObject.add_to_xmls_stack("word/#{target}")
        note = Note.parse(default_paragraph, default_character, target, header_reference.attribute('type').value, File.basename(target).sub('.xml', ''))
        page_properties.notes << note
        OOXMLDocumentObject.xmls_stack.pop
      end
      sect_pr.xpath('w:footerReference').each do |footer_reference|
        target = OOXMLDocumentObject.get_link_from_rels(footer_reference.attribute('id').value)
        OOXMLDocumentObject.add_to_xmls_stack("word/#{target}")
        note = Note.parse(default_paragraph, default_character, target, footer_reference.attribute('type').value, File.basename(target).sub('.xml', ''))
        page_properties.notes << note
        OOXMLDocumentObject.xmls_stack.pop
      end
      page_properties
    end
  end
end
