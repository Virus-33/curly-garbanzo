namespace Logic
{
    // Я не знал куда будет лучше прописать перечисления, чтобы они были доступны из любой части сборки.
    // Пусть пока будут здесь.
    enum GroupType
    {
        absentia,
        intramural
    }

    enum WorkType
    {
        lecture,
        practiceWork,
        lab,
        consultation,
        unmarkedExam,
        markedExam,
        courseProject,
        CGW,
        practice,
        graduationWork,
        GEC,
        magistracyLeader,
        internLeader
    }

    /// <summary>
    /// Класс, отвечающий за формирование отчёта и заполнение его полей на основе остальной информации.
    /// </summary>
    public class ReportShaper
    {
    }
}
