# English
## Roadmap and Future Development

This plan is a record of current issues and areas requiring research.

### Current Tasks:
- Finalizing the hierarchy: Complete the features from the CMFCBaseVisualManager base classes that did not make it into the 0.1 release.  
-  ~Advanced visual styles: Create classes for deeper visual styling (primarily layers between CMDIFrameWnd (and similar) and the user-inherited class for full non-client area support).~ (excluded. CNPUP_VisualManager will be used in the overall project - NOP PUP Controls Visual System for MFC)
- System components: Address scrollbars, system buttons, standard dialogs, and all others.
- Palette optimization: The current implementation works but requires a revision.
- Runtime customization: Implement customization interfaces at runtime (the capability exists now, but proper palettes must be established first).
- Overall optimization: Optimize everything possible (e.g., CDC pointers are currently "leased"; creating custom "memDC" every time is inefficient. The "if-else spaghetti" allowed for rapid implementation, but the final result... is impressive).
- Documentation: Create documentation, wiki, etc.

### Priorities
During the release preparation, many methods were refactored, and interesting capabilities of the CMFCBaseVisualManager family were discovered. Specifically, it was found that they allow access to the NC (non-client) area, though full control was not achieved as other directions were prioritized. Therefore, the primary focus now is researching all possibilities available within the CMFCVisualManager subclass (by overriding every possible method available throughout the CMFCBaseVisualManager lineage).

---

# Русский
## Планы и будущее развитие

Данный план — это фиксация текущих проблем и направлений, которые требуют исследования.   

### Текущие задачи:
- Добивка иерархии: Доделать то, что не вошло в релиз 0.1 из возможностей базовых классов семейства CMFCBaseVisualManager  
-  ~Создать классы реализующие более глубокую проработку визуальных стилей (в основном это должны быть прослойки между CMDIFrameWnd (и иже с ними) и пользовательским классом наследником, для полноценного подхвата неклиентской области).~ (исключено. CNPUP_VisualManager будет использован в общем проекте - NOP PUP Controls Visual System for MFC)
- Добраться до скроллбаров, решить вопрос с ними, до системных кнопок, стандартных диалогов и всех прочих.
- Оптимизация палитр: Текущая реализация работает, но требует пересмотра.
- Реализовать интерфейсы кастомизации на уровне Runtime (возможность есть хоть прямо сейчас, но.. но надо сначала сделать нормальные палитры).
- Оптимизация всего что можно оптимизировать (указатели на CDC даются в аренду, если создавать свой "memDC" - плохо будет. Спагети из if-else позволили сделать реализацию максимально быстро, но конечный результат... впечатляет).
- Составление документации, wiki...

### Приоритеты
В процессе подготовки релиза было переработано не мало методов и выявлено не мало интересных возможностей которые дают классы семейства CMFCBaseVisualManager, в частности было выяснено что они позволяют получить доступ к NC области, но добиться существенного результата по полному контролю над ней не получилось, так как предпочтение было отдано иным направлениям. Поэтому, в первую очередь, на текущий момент основным направлением является исследование всех возможностей доступных в рамках класса наследника CMFCVisualManager (с переопределением всех каких только можно методов имеющихся в распоряжении у всей родословной CMFCBaseVisualManager).  
