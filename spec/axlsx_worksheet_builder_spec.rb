# frozen_string_literal: true

require "spec_helper"

RSpec.describe AxlsxWorksheetBuilder do
  it "has a version number" do
    expect(AxlsxWorksheetBuilder::VERSION).not_to be_nil
  end

  context "with a simple object" do
    let(:data) do
      [{
        id: 1,
        name: "First name"
      }, {
        id: 2,
        name: "Second name"
      }]
    end

    let(:workbook) { Axlsx::Package.new.workbook }

    it "raises an error if the worksheet is not of type Axlsx::Worksheet" do
      expect do
        described_class.build("not a Axlsx::Worksheet", data) do |worksheet|
        end
      end.to raise_error(ArgumentError)
    end

    it "adds the corrects headers" do
      workbook.add_worksheet(name: "builder sheet") do |worksheet|
        described_class.build(worksheet, data) do |sheet|
          sheet.add_column("Id", property: :id)
          sheet.add_column("Name", property: :name)
        end
      end
      expect(workbook).to have_excel_cells(%w[Id Name]).in_row(0).in_sheet("builder sheet")
    end

    context "with properties" do
      before do
        workbook.add_worksheet(name: "builder sheet") do |worksheet|
          described_class.build(worksheet, data) do |sheet|
            sheet.add_column("Id", property: :id)
            sheet.add_column("Name", property: :name)
          end
        end
      end

      it "adds the corrects values" do
        aggregate_failures do
          expect(workbook).to have_excel_cells([1, "First name"]).in_row(1).in_sheet("builder sheet")
          expect(workbook).to have_excel_cells([2, "Second name"]).in_row(2).in_sheet("builder sheet")
        end
      end
    end

    context "with a block" do
      before do
        workbook.add_worksheet(name: "builder sheet") do |worksheet|
          described_class.build(worksheet, data) do |sheet|
            sheet.add_column("Id") do |v|
              v[:id]
            end
            sheet.add_column("Name") { |v| v[:name] }
          end
        end
      end

      it "adds the corrects values" do
        aggregate_failures do
          expect(workbook).to have_excel_cells([1, "First name"]).in_row(1).in_sheet("builder sheet")
          expect(workbook).to have_excel_cells([2, "Second name"]).in_row(2).in_sheet("builder sheet")
        end
      end
    end
  end

  context "with a 2 layers object" do
    let(:authors) do
      [{
        id: 1,
        name: "Martin",
        books: [{
          id: 2,
          title: "Ruby for Dummies"
        }]
      }]
    end
    let(:workbook) { Axlsx::Package.new.workbook }

    it "raises an error if the iterate through property is not an array" do
      authors = [{
        id: 1,
        name: "Martin",
        books: 2
      }]
      expect do
        workbook.add_worksheet(name: "builder sheet") do |worksheet|
          described_class.build(worksheet, authors) do |sheet|
            sheet.iterate_through_property(:books)
            sheet.add_column("Id", property: :id)
            sheet.add_column("Author") { |author| author[:name] }
            sheet.add_column("title", property: :title)
          end
        end
      end.to raise_error(NoMethodError, /undefined method [`|']each' for (2:|an instance of )Integer/)
    end

    it "has the correct values" do
      workbook.add_worksheet(name: "builder sheet") do |worksheet|
        described_class.build(worksheet, authors) do |sheet|
          sheet.iterate_through_property(:books)
          sheet.add_column("Id", property: :id)
          sheet.add_column("Author") { |author| author[:name] }
          sheet.add_column("title", property: :title)
        end
      end
      expect(workbook).to have_excel_cells([2, "Martin", "Ruby for Dummies"]).in_row(1).in_sheet("builder sheet")
    end
  end

  context "with a 2+ layers object" do
    let(:authors) do
      [{
        id: 1,
        name: "Martin",
        books: [{
          id: 2,
          title: "Ruby for Dummies",
          chapters: [{
            id: 3,
            title: "Chapter 1"
          }, {
            id: 3,
            title: "Chapter 2"
          }]
        }]
      }]
    end

    let(:workbook) { Axlsx::Package.new.workbook }

    before do
      workbook.add_worksheet(name: "builder sheet") do |worksheet|
        described_class.build(worksheet, authors) do |sheet|
          sheet.iterate_through_properties(:books, :chapters)
          sheet.add_column("Id") { |_author, book| book[:id] }
          sheet.add_column("Author") { |author| author[:name] }
          sheet.add_column("Book") { |_author, book| book[:title] }
          sheet.add_column("Chapter", property: :title)
        end
      end
    end

    it "has the correct values" do
      aggregate_failures do
        expect(workbook).to have_excel_cells([2, "Martin", "Ruby for Dummies",
                                              "Chapter 1"]).in_row(1).in_sheet("builder sheet")
        expect(workbook).to have_excel_cells([2, "Martin", "Ruby for Dummies",
                                              "Chapter 2"]).in_row(2).in_sheet("builder sheet")
      end
    end
  end

  it "allows specifying column types" do
    data = [{
      id: "001",
      name: "Martin"
    }]
    workbook = Axlsx::Package.new.workbook
    workbook.add_worksheet(name: "builder sheet") do |worksheet|
      described_class.build(worksheet, data) do |sheet|
        sheet.add_column("Id", property: :id, type: :string)
        sheet.add_column("Name", property: :name)
      end
    end
    expect(workbook).to have_excel_cells(%w[001 Martin]).in_row(1).in_sheet("builder sheet")
  end
end
