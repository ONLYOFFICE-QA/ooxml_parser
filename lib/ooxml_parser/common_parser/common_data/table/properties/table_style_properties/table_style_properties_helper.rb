module OoxmlParser
  # Helper method for working with TableStyleProperties
  module TableStylePropertiesHelper
    table_styles_names_hash = { first_column: :firstCol,
                                last_column: :lastCol,
                                whole_table: :wholeTbl,
                                banding_1_horizontal:  :band1Horz,
                                banding_2_horizontal: :band2Horz,
                                banding_1_vertical: :band1Vert,
                                banding_2_vertical: :band2Vert,
                                southeast_cell: :seCell,
                                southwest_cell: :swCell,
                                northeast_cell: :neCell,
                                northwest_cell: :nwCell,
                                first_row: :firstRow,
                                last_row: :lastRow }.freeze

    table_styles_names_hash.keys.each do |table_style_name|
      define_method(table_style_name) do
        @table_style_properties_list.each do |table_style|
          return table_style if table_style.type == table_styles_names_hash[table_style_name]
        end
        nil
      end
    end
  end
end
