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
        page_properties.size = Size.new
        page_properties.size.orientation = pg_sz.attribute('orient').value.to_sym unless pg_sz.attribute('orient').nil?
        page_properties.size.height = pg_sz.attribute('h').value
        page_properties.size.width = pg_sz.attribute('w').value
      end
      sect_pr.xpath('w:pgBorders').each do |pg_borders|
        page_borders = Borders.new
        unless pg_borders.attribute('display').nil?
          page_borders.display = pg_borders.attribute('display').value
        end
        unless pg_borders.attribute('offsetFrom').nil?
          page_borders.offset_from = pg_borders.attribute('offsetFrom').value
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
        page_properties.page_border = page_borders
      end
      sect_pr.xpath('w:type').each do |type|
        page_properties.type = type.attribute('val').value
      end
      sect_pr.xpath('w:pgMar').each do |pg_mar|
        page_properties.margins = PageMargins.new
        page_properties.margins.footer = (pg_mar.attribute('footer').value.to_f / 566.9).round(2) unless pg_mar.attribute('footer').nil?
        page_properties.margins.gutter = (pg_mar.attribute('gutter').value.to_f / 566.9).round(2)
        page_properties.margins.header = (pg_mar.attribute('header').value.to_f / 566.9).round(2) unless pg_mar.attribute('header').nil?
        page_properties.margins.bottom = (pg_mar.attribute('bottom').value.to_f / 566.9).round(2)
        page_properties.margins.left = (pg_mar.attribute('left').value.to_f / 566.9).round(2)
        page_properties.margins.right = (pg_mar.attribute('right').value.to_f / 566.9).round(2)
        page_properties.margins.top = (pg_mar.attribute('top').value.to_f / 566.9).round(2)
      end
      sect_pr.xpath('w:pgNumType').each do |pg_num_type|
        page_properties.num_type = pg_num_type.attribute('fmt').value
      end
      sect_pr.xpath('w:formProt').each do |form_prot|
        page_properties.form_prot = form_prot.attribute('val').value
      end
      sect_pr.xpath('w:textDirection').each do |text_direction|
        page_properties.text_direction = text_direction.attribute('val').value
      end
      sect_pr.xpath('w:textDirection').each do |text_direction|
        page_properties.document_grid = DocumentGrid.new
        page_properties.document_grid.char_space = text_direction.attribute('charSpace').value
        page_properties.document_grid.line_pitch = text_direction.attribute('linePitch').value
        page_properties.document_grid.type = text_direction.attribute('type').value
      end
      sect_pr.xpath('w:cols').each do |columns|
        columns_count = 1
        columns_count = columns.attribute('num').value.to_i unless columns.attribute('num').nil?
        page_properties.columns = Columns.new(columns_count)
        page_properties.columns.separator = columns.attribute('sep').value unless columns.attribute('sep').nil?
        i = 0
        columns.xpath('w:col').each do |col|
          width = col.attribute('w').value
          width = StringHelper.round(width.to_f / 566.9, 2) unless width.nil?
          space = col.attribute('space').value
          space = StringHelper.round(space.to_f / 566.9, 2) unless space.nil?
          page_properties.columns.columns[i] = Column.new(width, space)
          i += 1
        end
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
