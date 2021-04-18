import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelReaderComponent } from './excel-reader.component';

describe('ExcelReaderComponent', () => {
  let component: ExcelReaderComponent;
  let fixture: ComponentFixture<ExcelReaderComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExcelReaderComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(ExcelReaderComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
